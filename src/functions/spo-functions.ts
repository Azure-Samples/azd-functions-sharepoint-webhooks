import { app } from "@azure/functions";
import { getWebTitle, setListItem } from "../spo/spoapp.js"

app.http('getWebTitle', { methods: ['GET'], authLevel: 'function', handler: getWebTitle, route: 'sharepoint/getWebTitle' });
app.http('setListItem', { methods: ['POST'], authLevel: 'function', handler: setListItem, route: 'sharepoint/setListItem' });
