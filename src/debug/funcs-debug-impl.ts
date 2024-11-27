import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import "@pnp/sp/items/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/subscriptions/index.js";
import "@pnp/sp/webs/index.js";
import { CommonConfig, safeWait } from "../utils/common.js";
import { handleError } from "../utils/loggingHandler.js";
import { getSharePointSiteInfo, getSpAccessToken, getSPFI } from "../utils/spAuthentication.js";

export async function getAccessToken(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const token = await getSpAccessToken(tenantPrefix || CommonConfig.TenantPrefix);
        return { status: 200, jsonBody: token };
    }
    catch (error: unknown) {
        const errMessage = await handleError(error, context, `Unexpected error whhile executing the function: `);
        return { status: 400, body: errMessage };
    }
};

export async function showWeb(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        let result: any, error: any;
        [result, error] = await safeWait(sp.web());
        if (error) {
            const errMessage = await handleError(error, context, `Could not get web for tenantPrefix "${sharePointSite.tenantPrefix}" and site "${sharePointSite.siteRelativePath}"`);
            return { status: 400, jsonBody: { status: 'error', message: errMessage } };
        }
        return { status: 200, jsonBody: result };
    }
    catch (error: unknown) {
        const errMessage = await handleError(error, context, `Unexpected error whhile executing the function: `);
        return { status: 400, jsonBody: { status: 'error', message: errMessage } };
    }
};
