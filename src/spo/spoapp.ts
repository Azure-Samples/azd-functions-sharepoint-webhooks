import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import "@pnp/sp/items/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/webs/index.js";
import { getSharePointSiteInfo, safeWait } from "../common.js";
import { getSPFI } from "../spAuthentication.js";
import { Logger, LogLevel } from "@pnp/logging";
import { handleError } from "../loggingHandler.js";

export async function getWebTitle(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const siteRelativePath = request.query.get('siteRelativePath') || undefined;
    const tenantPrefix = request.query.get('tenantPrefix') || undefined;

    const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
    const sp = getSPFI(sharePointSite);
    let error: any, webData: any;
    [webData, error] = await safeWait(sp.web());
    if (error) {
        const errMessage = await handleError(error, context, `Could not get the SharePoint web details: `);
        return { status: 400, body: errMessage };
    }

    const jsonBody = { title: webData.Title };
    Logger.log({ data: context, message: `Connection to the SharePoint web OK: "${JSON.stringify(jsonBody)}"`, level: LogLevel.Info });
    return {
        jsonBody: jsonBody
    };
};

export async function setListItem(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const siteRelativePath = request.query.get('siteRelativePath') || undefined;
        const tenantPrefix = request.query.get('tenantPrefix') || undefined;
        const listTitle = request.query.get('listTitle');
        const itemTitle = request.query.get('itemTitle');
        const itemValue: string = request.query.get('itemValue') || new Date().toISOString();
        if (!listTitle) { return { status: 400, body: 'Value listTitle is required' }; }

        const sharePointSite = getSharePointSiteInfo(tenantPrefix, siteRelativePath);
        const sp = getSPFI(sharePointSite);
        const list = sp.web.lists.getByTitle(listTitle);
        if (!list) {
            const errMessage = `List '${listTitle}' was not found`;
            Logger.log({ data: context, message: errMessage, level: LogLevel.Error });
            return { status: 400, body: errMessage };
        }

        let items: any[], item: any, error: any;
        [items, error] = await safeWait(list.items.select("Id", "Title").filter(`Title eq '${itemTitle}'`)());
        if (error) {
            const errMessage = await handleError(error, context, `Unexpected error whhile getting items from list "${itemTitle}": `);
            return { status: 400, body: errMessage };
        }

        if (items.length > 0) {
            item = items[0]
        }

        let jsonBody;
        if (item) {
            Logger.log({ data: context, message: `Updating item '${item.Title}' in list '${listTitle}' with value '${itemValue}'...`, level: LogLevel.Info });
            await sp.web.lists.getByTitle(listTitle).items.getById(item.Id).update({
                Description: itemValue,
            });
            // https://pnp.github.io/pnpjs/transition-guide/#addupdate-methods-no-longer-returning-data-and-a-queryable-instance
            // "update events return 204, which would translate into a return type of void. In that case you will have to adjust your code to make a second call "
            let updatedItem = await sp.web.lists.getByTitle(listTitle).items.getById(item.Id)();
            Logger.log({ data: context, message: `Updated item: '${JSON.stringify(updatedItem)}'.`, level: LogLevel.Info });
            jsonBody = updatedItem;
            jsonBody.Operation = "Update";
        } else {
            Logger.log({ data: context, message: `Adding item '${itemTitle}' in list '${listTitle}' with value '${itemValue}'...`, level: LogLevel.Info });
            const addedItem = await sp.web.lists.getByTitle(listTitle).items.add({
                Title: itemTitle,
                Description: itemValue
            });
            Logger.log({ data: context, message: `Added item: '${JSON.stringify(addedItem)}'.`, level: LogLevel.Info });
            jsonBody = addedItem;
            jsonBody.Operation = "Add";
        }
        return { jsonBody: jsonBody };
    }
    catch (error: unknown) {
        const errMessage = await handleError(error, context, `Unexpected error whhile running function: `);
        return { status: 400, body: errMessage };
    }
};
