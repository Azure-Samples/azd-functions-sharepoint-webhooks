import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { dateAdd } from "@pnp/core";
import { IChangeQuery, SPFI } from "@pnp/sp";
import "@pnp/sp/items/index.js";
import "@pnp/sp/lists/index.js";
import { IListEnsureResult } from "@pnp/sp/lists/types.js";
import "@pnp/sp/subscriptions/index.js";
import "@pnp/sp/webs/index.js";
import { CommonConfig, GetChangeTokenTicks, ISharePointWeebhookEvent, ISharePointWeebhookEventValue, ISubscriptionResponse, safeWait, WebhookChangeType } from "../utils/common.js";
import { logError, logMessage } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSPFI } from "../utils/spAuthentication.js";

async function registerWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const listTitle = request.query.get('listTitle');
        const notificationUrl = request.query.get('notificationUrl');

        if (!listTitle || !notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        const expiryDate: Date = dateAdd(new Date(), "day", 180) as Date; // Set the expiry date to 180 days from now, which is the maximum allowed for the webhook expiry date.
        let result: any, error: any;
        [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions.add(notificationUrl, expiryDate.toISOString()));
        if (error) {
            const errorDetails = await logError(context, error, `Could not register webhook '${notificationUrl}' in list '${listTitle}'`);
            return { status: errorDetails.httpStatus, jsonBody: errorDetails };
        }
        logMessage(context, `Attempted to register webhook '${notificationUrl}' to list '${listTitle}' with expiry date '${expiryDate.toISOString()}'. Result: ${JSON.stringify(result)}`);
        return { status: 200, jsonBody: result };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

async function wehhookService(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const validationtoken = request.query.get('validationtoken');
        if (validationtoken) {
            logMessage(context, `Validated webhook registration with validation token: ${validationtoken}`);
            return { status: 200, headers: { 'Content-Type': 'text/plain' }, body: validationtoken };
        }

        const body: ISharePointWeebhookEvent = await request.json() as ISharePointWeebhookEvent;
        const sharePointSite = getSharePointSiteInfo();
        const sp = getSPFI(sharePointSite);

        let webhookEventsChanges: Promise<any>[] = [];
        for (let i: number = 0; i < body.value.length; i++) {
            webhookEventsChanges.push(getListChanges(context, sp, body.value[i]));
        }
        const webhookEventsResult: string[] = await Promise.all(webhookEventsChanges);

        // Log the webhook notification in the history list
        let webhookHistoryListEnsureResult: IListEnsureResult, error: any;
        [webhookHistoryListEnsureResult, error] = await safeWait(sp.web.lists.ensure(CommonConfig.WebhookHistoryListTitle));
        if (error) {
            const errorDetails = await logError(context, error, `Could not ensure that list '${CommonConfig.WebhookHistoryListTitle}' exists`);
            return { status: errorDetails.httpStatus };
        }
        if (webhookHistoryListEnsureResult.created === true) {
            logMessage(context, `List '${CommonConfig.WebhookHistoryListTitle}' (to log the webhook notifications) did not exist and was just created.`);
        }

        for (let i: number = 0; i < webhookEventsResult.length; i++) {
            let result: any;
            [result, error] = await safeWait(sp.web.lists.getByTitle(CommonConfig.WebhookHistoryListTitle).items.add({
                Title: webhookEventsResult[i],
            }));
            if (error) {
                const errorDetails = await logError(context, error, `Could not add an item to the list '${CommonConfig.WebhookHistoryListTitle}'`);
                return { status: errorDetails.httpStatus };
            }
        }
        logMessage(context, `Processed the nofification from SharePoint, which contained ${body.value.length} event(s).`);
        return { status: 200 };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus };
    }
};

async function getListChanges(context: InvocationContext, sp: SPFI, webhookNotification: ISharePointWeebhookEventValue): Promise<string> {
    // Get all changes since some minutes ago
    const listId = webhookNotification.resource;
    const minutesFromNow: number = CommonConfig.WebhookChangesMinutesAgo;
    const changeStartTicks = GetChangeTokenTicks(minutesFromNow);
    const changeTokenStart = `1;3;${listId};${changeStartTicks};-1`;
    const changeQuery: IChangeQuery = {
        ChangeTokenStart: { StringValue: changeTokenStart },
        ChangeTokenEnd: undefined,
        Item: true,
        Add: true,
        DeleteObject: true,
        Update: true,
        Rename: true,
        Restore: true,
    };

    let changes: any[], error: any;
    [changes, error] = await safeWait(sp.web.lists.getById(listId).getChanges(changeQuery));
    if (error) {
        const errorDetails = await logError(context, error, `Could not get the changes for list '${listId}' in '${webhookNotification.siteUrl}'.`);
        return errorDetails.message;
    }

    const numberOfAdds: number = changes.filter(c => c.ChangeType === WebhookChangeType.Add)?.length || 0;
    const numberOfUpdates: number = changes.filter(c => c.ChangeType === WebhookChangeType.Update)?.length || 0;
    const numberOfDeletes: number = changes.filter(c => c.ChangeType === WebhookChangeType.Delete)?.length || 0;
    const message = `${changes.length} change(s) found in list '${listId}' in '${webhookNotification.siteUrl}' since ${minutesFromNow} minutes, including ${numberOfAdds} add(s), ${numberOfUpdates} update(s) and ${numberOfDeletes} delete(s).`;
    return message;
}

async function listWehhooks(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const listTitle = request.query.get('listTitle');

        if (!listTitle) { return { status: 400, body: `Required parameters are missing.` }; }

        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        let result: any, error: any;
        [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions());
        if (error) {
            const errorDetails = await logError(context, error, `Could not list webhook for web '${sharePointSite.siteRelativePath}' and list '${listTitle}'`);
            return { status: errorDetails.httpStatus, jsonBody: errorDetails };
        };
        logMessage(context, `Webhooks registered on web '${sharePointSite.siteRelativePath}' and list '${listTitle}': ${JSON.stringify(result)}`);
        return { status: 200, jsonBody: result };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

async function showWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const notificationUrl = request.query.get('notificationUrl');
        if (!notificationUrl) { return { status: 400, body: `Required parameters are missing.` }; }

        const webhooksResponse = await listWehhooks(request, context);
        if (!webhooksResponse || !webhooksResponse.jsonBody) { return { status: 200, jsonBody: {} }; }
        if (webhooksResponse.status !== 200) { return webhooksResponse; }
        const webhooks: ISubscriptionResponse[] = webhooksResponse.jsonBody;
        const webhook = webhooks.find((element) => element.notificationUrl === notificationUrl);
        return { status: 200, jsonBody: webhook ? webhook : {} };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: { status: 'error', message: errorDetails } };
    }
};

async function removeWehhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const listTitle = request.query.get('listTitle');
        const webhookId = request.query.get('webhookId');

        if (!listTitle || !webhookId) { return { status: 400, body: `Required parameters are missing.` }; }

        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        let result: any, error: any;
        [result, error] = await safeWait(sp.web.lists.getByTitle(listTitle).subscriptions.getById(webhookId).delete());
        if (error) {
            const errorDetails = await logError(context, error, `Could not delete webhook '${webhookId}' for web '${sharePointSite.siteRelativePath}' and list '${listTitle}'`);
            return { status: errorDetails.httpStatus, jsonBody: errorDetails };
        }
        logMessage(context, `Deleted webhook '${webhookId}' registered on web '${sharePointSite.siteRelativePath}' and list '${listTitle}'.`);
        return { status: 204 };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: { status: 'error', message: errorDetails } };
    }
};


app.http('webhooks-register', { methods: ['POST'], authLevel: 'function', handler: registerWebhook, route: 'webhooks/register' });
app.http('webhooks-service', { methods: ['POST'], authLevel: 'function', handler: wehhookService, route: 'webhooks/service' });
app.http('webhooks-list', { methods: ['GET'], authLevel: 'function', handler: listWehhooks, route: 'webhooks/list' });
app.http('webhooks-remove', { methods: ['POST'], authLevel: 'function', handler: removeWehhook, route: 'webhooks/remove' });
app.http('webhooks-show', { methods: ['GET'], authLevel: 'function', handler: showWehhook, route: 'webhooks/show' });
