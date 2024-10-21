import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/items/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/webs/index.js";
import { safeWait } from "../common.js";
import { getSPFI } from "../spAuthentication.js";


export async function getWebTitle(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const sp: SPFI = getSPFI();
    let body: string;
    try {
        // let[webData, error] = [{ Title: "webdataTitle" }, undefined];
        let [webData, error] = await safeWait(sp.web());
        if (error) {
            context.error(`Unexpected error whhile trying to connect to SPO: ${error}`);
            body = error;
        } else {
            context.log(`connection to SPO site successful: Title: ${webData.Title}`);
            body = webData.Title;
        }
    }
    catch (ex) {
        context.error(ex);
        body = ex as string;
    }
    return { body: body };
};

export async function setListItem(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const sp: SPFI = getSPFI();
    let body: string = "";
    let currentMessage: string = "";
    try {
        const listTitle = request.query.get('listTitle');
        const itemTitle = request.query.get('itemTitle');
        let itemValue = request.query.get('itemValue');

        if (!listTitle) { return { status: 400, body: 'Value listTitle is required' }; }
        if (!itemTitle) { return { status: 400, body: 'Value itemTitle is required' }; }

        let error: any;
        const list = sp.web.lists.getByTitle(listTitle);
        if (!list) {
            currentMessage = `List '${listTitle}' was not found`;
            context.error(currentMessage);
            body += `\n${currentMessage}`;
            return { body: body };
        }

        let items: any[];
        [items, error] = await safeWait(list.items.select("Id", "Title").filter(`Title eq '${itemTitle}'`)());
        if (error) {
            currentMessage = `Unexpected error whhile trying to connect to SPO: ${error}`;
            context.error(currentMessage);
            body += `\n${currentMessage}`;
            return { body: body };
        }
        
        let item: any;
        if (items.length > 0) {
            item = items[0]
            currentMessage = `Found item '${item.Title}' in list '${listTitle}'`;
            context.log(currentMessage);
            body += `\n${currentMessage}`;
        }

        if (!itemValue) {
            itemValue = new Date().toISOString();
        }

        if (item) {
            currentMessage = `Updating item '${item.Title}' in list '${listTitle}' with value '${itemValue}'...`;
            context.log(currentMessage);
            body += `\n${currentMessage}`;
            await sp.web.lists.getByTitle(listTitle).items.getById(item.Id).update({
                Description: itemValue,
            });
            // https://pnp.github.io/pnpjs/transition-guide/#addupdate-methods-no-longer-returning-data-and-a-queryable-instance
            // "update events return 204, which would translate into a return type of void. In that case you will have to adjust your code to make a second call "
            let updatedItem = await sp.web.lists.getByTitle(listTitle).items.getById(item.Id)();
            currentMessage = `Updated item: '${JSON.stringify(updatedItem)}'.`;
            context.log(currentMessage);
            body += `\n${currentMessage}`;
        } else {
            currentMessage = `Adding item '${itemTitle}' in list '${listTitle}' with value '${itemValue}'...`;
            context.log(currentMessage);
            body += `\n${currentMessage}`;

            const addedItem = await sp.web.lists.getByTitle(listTitle).items.add({
                Title: itemTitle,
                Description: itemValue
            });
            currentMessage = `Added item: '${JSON.stringify(addedItem)}'.`;
            context.log(currentMessage);
            body += `\n${currentMessage}`;
        }
    }
    catch (ex) {
        context.error(ex);
        body += ex as string;
    }
    return { body: body };
};
