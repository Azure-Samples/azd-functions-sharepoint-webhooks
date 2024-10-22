import { Logger, LogLevel } from "@pnp/logging";
import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { dateAdd } from "@pnp/core";
import "@pnp/sp/subscriptions/index.js";
import "@pnp/sp/webs/index.js";
import { ISubscriptionResponse, safeWait } from "../common.js";
import { getSPFI } from "../spAuthentication.js";
import { handleError } from "../loggingHandler.js";

export async function registerWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath');
    const tenantPrefix = request.query.get('tenantPrefix');
    const listTitle = request.query.get('listTitle');
    const notificationUrl = request.query.get('notificationUrl');

    if (!listTitle || !notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

    let sharePointSite = undefined;
    if (siteRelativePath && tenantPrefix) {
        sharePointSite = { tenantPrefix: tenantPrefix, siteRelativePath: siteRelativePath };
    }

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

    const body = await request.json();
    let message = `Received webhook notification: ${JSON.stringify(body)}`;
    Logger.log({ data: context, message: message, level: LogLevel.Info });
    return { body: message };
};

export async function listRegisteredWehhooks(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath');
    const tenantPrefix = request.query.get('tenantPrefix');
    const listTitle = request.query.get('listTitle');

    if (!listTitle) { return { status: 400, body: `Required parameters are missing.` }; }

    let sharePointSite = undefined;
    if (siteRelativePath && tenantPrefix) {
        sharePointSite = { tenantPrefix: tenantPrefix, siteRelativePath: siteRelativePath };
    }

    const sp = getSPFI(sharePointSite);
    let result: any, error: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions());
    if (error) {
        return { status: 400, body: await handleError(error, context, `Could not list webhook for web "${siteRelativePath}" and list "${listTitle}": "${error}"`) };
    }
    Logger.log({ data: context, message: `Webhooks registered on web "${siteRelativePath}" and list "${listTitle}": ${JSON.stringify(result)}`, level: LogLevel.Info });
    return { body: `{ "webhooks": ${JSON.stringify(result)} }` };
};

export async function showRegisteredWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const notificationUrl = request.query.get('notificationUrl');
    if (!notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

    const webhooks = await listRegisteredWehhooks(request, context);
    if (!webhooks || !webhooks.body) { return { status: 200, body: `No webhook found.` }; }
    const webhooksBody = JSON.parse(webhooks.body.toString());
    const webhooksJson: ISubscriptionResponse[] = webhooksBody.webhooks;
    const webhook = webhooksJson.find((element) => element.notificationUrl === notificationUrl);
    return { body: JSON.stringify(webhook) };
};

export async function removeRegisteredWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath');
    const tenantPrefix = request.query.get('tenantPrefix');
    const listTitle = request.query.get('listTitle');
    const webhookId = request.query.get('webhookId');

    if (!listTitle || !webhookId) { return { status: 400, body: `Required parameters are missing.` }; }

    let sharePointSite = undefined;
    if (siteRelativePath && tenantPrefix) {
        sharePointSite = { tenantPrefix: tenantPrefix, siteRelativePath: siteRelativePath };
    }

    const sp = getSPFI(sharePointSite);
    let result: any, error: any;
    [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions.getById(webhookId).delete());
    if (error) {
        return { status: 400, body: await handleError(error, context, `Could not delete webhook "${webhookId}" for web "${siteRelativePath}" and list "${listTitle}": "${error}"`) };
    }
    Logger.log({ data: context, message: `Deleted webhook "${webhookId}" registered on web "${siteRelativePath}" and list "${listTitle}".`, level: LogLevel.Info });
    return { status: 204 };
};
