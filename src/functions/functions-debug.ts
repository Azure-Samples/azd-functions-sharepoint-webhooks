import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { CommonConfig, safeWait } from "../utils/common.js";
import { logError } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSpAccessToken, getSPFI } from "../utils/spAuthentication.js";

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

app.http('debug-getAccessToken', { methods: ['GET'], authLevel: 'admin', handler: getAccessToken, route: 'debug/getAccessToken' });
app.http('debug-getWeb', { methods: ['GET'], authLevel: 'function', handler: getWeb, route: 'debug/getWeb' });
