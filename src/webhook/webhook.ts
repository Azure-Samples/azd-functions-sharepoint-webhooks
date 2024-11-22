import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { dateAdd } from "@pnp/core";
import { Logger, LogLevel } from "@pnp/logging";
import "@pnp/sp/subscriptions/index.js";
import "@pnp/sp/webs/index.js";
import { CommonConfig, ISharePointWeebhookEvent, ISubscriptionResponse, safeWait } from "../utils/common.js";
import { handleError } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSPFI } from "../utils/spAuthentication.js";
import { IListEnsureResult } from "@pnp/sp/lists/types.js";

export async function registerWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath') || undefined;
    const tenantPrefix = request.query.get('tenantPrefix') || undefined;
    const listTitle = request.query.get('listTitle');
    const notificationUrl = request.query.get('notificationUrl');

    if (!listTitle || !notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

    const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
    const sp = getSPFI(sharePointSite);
    const expiryDate: Date = dateAdd(new Date(), "day", 180) as Date; // Set the expiry date to 180 days from now, which is the maximum allowed for the webhook expiry date.
    let message: string, result: any, error: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions.add(notificationUrl, expiryDate.toISOString()));
    if (error) {
        message = await handleError(error, context, `Could not register webhook "${notificationUrl}" in list "${listTitle}": `);
        return { status: 400, body: message };
    }
    Logger.log({ data: context, message: `Attempted to register webhook "${notificationUrl}" to list "${listTitle}" with expiry date "${expiryDate.toISOString()}". Result: ${JSON.stringify(result)}`, level: LogLevel.Info });
    return { body: JSON.stringify(result) };
};

export async function wehhookService(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const validationtoken = request.query.get('validationtoken');
    if (validationtoken) {
        Logger.log({ data: context, message: `Validated webhook registration with validation token: ${validationtoken}`, level: LogLevel.Info });
        return { headers: { 'Content-Type': 'text/plain' }, body: validationtoken };
    }

    const body: ISharePointWeebhookEvent = await request.json() as ISharePointWeebhookEvent;
    const message = `Received webhook notification: ${body.value.length} events for resource \"${body.value[0].resource}\" on site \"${body.value[0].siteUrl}\"`;
    Logger.log({ data: context, message: message, level: LogLevel.Info });

    const sharePointSite = getSharePointSiteInfo();
    const sp = getSPFI(sharePointSite);
    let webhookHistoryListEnsureResult: IListEnsureResult, error: any;
    [webhookHistoryListEnsureResult, error] = await safeWait(sp.web.lists.ensure(CommonConfig.WebhookHistoryListTitle));
    if (error) {
        await handleError(error, context, `Could not ensure that list "${CommonConfig.WebhookHistoryListTitle}" exists: `);
        return { body: '' };
    }
    if (webhookHistoryListEnsureResult.created === true) {
        let message = `List "${CommonConfig.WebhookHistoryListTitle}" (to log the webhook notifications) did not exist and was just created.`;
        Logger.log({ data: context, message: message, level: LogLevel.Info });
    }
    let result: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(CommonConfig.WebhookHistoryListTitle).items.add({
        Title: message
    }));
    if (error) {
        await handleError(error, context, `Could not add an item to the list "${CommonConfig.WebhookHistoryListTitle}": `);
        return { body: '' };
    }

    return { body: '' };
};

export async function listWehhooks(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath') || undefined;
    const tenantPrefix = request.query.get('tenantPrefix') || undefined;
    const listTitle = request.query.get('listTitle');

    if (!listTitle) { return { status: 400, body: `Required parameters are missing.` }; }

    const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
    const sp = getSPFI(sharePointSite);
    let result: any, error: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions());
    if (error) {
        return { status: 400, body: await handleError(error, context, `Could not list webhook for web "${sharePointSite.siteRelativePath}" and list "${listTitle}"`) };
    }
    Logger.log({ data: context, message: `Webhooks registered on web "${sharePointSite.siteRelativePath}" and list "${listTitle}": ${JSON.stringify(result)}`, level: LogLevel.Info });
    return { body: `{ "webhooks": ${JSON.stringify(result)} }` };
};

export async function showWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const notificationUrl = request.query.get('notificationUrl');
    if (!notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

    const webhooks = await listWehhooks(request, context);
    if (!webhooks || !webhooks.body) { return { status: 200, body: `No webhook found.` }; }
    const webhooksBody = JSON.parse(webhooks.body.toString());
    const webhooksJson: ISubscriptionResponse[] = webhooksBody.webhooks;
    const webhook = webhooksJson.find((element) => element.notificationUrl === notificationUrl);
    return { body: JSON.stringify(webhook) };
};

export async function removeWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath') || undefined;
    const tenantPrefix = request.query.get('tenantPrefix') || undefined;
    const listTitle = request.query.get('listTitle');
    const webhookId = request.query.get('webhookId');

    if (!listTitle || !webhookId) { return { status: 400, body: `Required parameters are missing.` }; }

    const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
    const sp = getSPFI(sharePointSite);
    let result: any, error: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions.getById(webhookId).delete());
    if (error) {
        return { status: 400, body: await handleError(error, context, `Could not delete webhook "${webhookId}" for web "${sharePointSite.siteRelativePath}" and list "${listTitle}"`) };
    }
    Logger.log({ data: context, message: `Deleted webhook "${webhookId}" registered on web "${sharePointSite.siteRelativePath}" and list "${listTitle}".`, level: LogLevel.Info });
    return { status: 204 };
};
