import { app } from "@azure/functions";
import { getAccessToken, showWeb } from "../debug/funcs-debug-impl.js";

app.http('debug-getAccessToken', { methods: ['GET'], authLevel: 'function', handler: getAccessToken, route: 'debug/getAccessToken' });
app.http('debug-showWeb', { methods: ['GET'], authLevel: 'function', handler: showWeb, route: 'debug/showWeb' });
