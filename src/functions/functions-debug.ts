import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { IChangeQuery } from "@pnp/sp";
import { CommonConfig, GetChangeTokenTicks, safeWait, WebhookChangeType } from "../utils/common.js";
import { logError, logMessage } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSpAccessToken, getSPFI } from "../utils/spAuthentication.js";
import { getListChanges } from "./functions-webhooks.js";

async function getAccessToken(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const tenantPrefix = request.query.get('tenantPrefix') || CommonConfig.TenantPrefix;
    try {
        const token = await getSpAccessToken(tenantPrefix);
        let result: any = {
            userAssignedManagedIdentityClientId: CommonConfig.UserAssignedManagedIdentityClientId,
            tenantPrefix: tenantPrefix,
            sharePointDomain: CommonConfig.SharePointDomain,
            token: token,
        };
        return { status: 200, jsonBody: result };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

export async function getWeb(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        let result: any, error: any;
        [result, error] = await safeWait(sp.web());
        if (error) {
            const errorDetails = await logError(context, error, `Could not get web for tenantPrefix '${sharePointSite.tenantPrefix}' and site '${sharePointSite.siteRelativePath}'`);
            return { status: errorDetails.httpStatus, jsonBody: errorDetails };
        }
        return { status: 200, jsonBody: result };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

export async function getListChangesFunc(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const listId = request.query.get('listId');
        const minutesFromNow: number = Number(request.query.get('minutesFromNow')) || CommonConfig.MinutesFromNowForChanges;
        if (!listId) { return { status: 400, body: `Required parameters are missing.` }; }

        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        
        // Get all changes since some minutes ago
        const webhookListId = listId;
        const changeStartTicks = GetChangeTokenTicks(minutesFromNow);
        const changeTokenStart = `1;3;${webhookListId};${changeStartTicks};-1`;
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
        const changes: any[] = await sp.web.lists.getById(webhookListId).getChanges(changeQuery);
        const numberOfAdds: number = changes.filter(c => c.ChangeType === WebhookChangeType.Add)?.length || 0;
        const numberOfUpdates: number = changes.filter(c => c.ChangeType === WebhookChangeType.Update)?.length || 0;
        const numberOfDeletes: number = changes.filter(c => c.ChangeType === WebhookChangeType.Delete)?.length || 0;
        const message = logMessage(context, `${changes.length} change(s) found in list '${webhookListId}' since ${minutesFromNow} minutes, including ${numberOfAdds} add(s), ${numberOfUpdates} update(s) and ${numberOfDeletes} delete(s).`);
        // return { status: 200, body: JSON.stringify(changes, ['ChangeType', 'ItemId']) };
        return { status: 200, jsonBody: message };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

app.http('debug-getAccessToken', { methods: ['GET'], authLevel: 'admin', handler: getAccessToken, route: 'debug/getAccessToken' });
app.http('debug-getWeb', { methods: ['GET'], authLevel: 'function', handler: getWeb, route: 'debug/getWeb' });
app.http('debug-getChanges', { methods: ['GET'], authLevel: 'function', handler: getListChangesFunc, route: 'debug/getChanges' });
