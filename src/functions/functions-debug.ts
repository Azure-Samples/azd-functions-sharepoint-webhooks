import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { CommonConfig, safeWait } from "../utils/common.js";
import { logError } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSpAccessToken, getSPFI } from "../utils/spAuthentication.js";
import { IChangeQuery } from "@pnp/sp";
import { dateAdd } from "@pnp/core/util.js";

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

export async function getChanges(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const listId = request.query.get('listId');
        if (!listId) { return { status: 400, body: `Required parameters are missing.` }; }


        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        // build the changeQuery object to get any change since 5 minutes ago
        const now = new Date();
        const changeStart = ((now.getTime() * 10000 - 5 * 60_000 * 10000) + 621355968000000000)
        const webhookListId = listId;
        const changeTokenStart = `1;3;${webhookListId};${changeStart};-1`;
        const changeQuery: IChangeQuery = {
            ChangeTokenStart: { StringValue: changeTokenStart },
            // ChangeTokenStart: undefined,
            ChangeTokenEnd: undefined,
            Add: true,
            DeleteObject: true,
            Rename: true,
            Restore: true,
            Item: true,
        };
        const changes: any[] = await sp.web.lists.getById(webhookListId).getChanges(changeQuery);
        return { status: 200, body: JSON.stringify(changes, ['ChangeType', 'ItemId']) };
    }
    catch (error: unknown) {
        const errorDetails = await logError(context, error, context.functionName);
        return { status: errorDetails.httpStatus, jsonBody: errorDetails };
    }
};

app.http('debug-getAccessToken', { methods: ['GET'], authLevel: 'admin', handler: getAccessToken, route: 'debug/getAccessToken' });
app.http('debug-getWeb', { methods: ['GET'], authLevel: 'function', handler: getWeb, route: 'debug/getWeb' });
app.http('debug-getChanges', { methods: ['GET'], authLevel: 'function', handler: getChanges, route: 'debug/getChanges' });
