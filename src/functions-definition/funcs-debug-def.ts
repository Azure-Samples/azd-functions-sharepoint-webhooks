import { app } from "@azure/functions";
import { getAccessToken, getWeb } from "../debug/funcs-debug-impl.js";

app.http('debug-getAccessToken', { methods: ['GET'], authLevel: 'admin', handler: getAccessToken, route: 'debug/getAccessToken' });
app.http('debug-getWeb', { methods: ['GET'], authLevel: 'function', handler: getWeb, route: 'debug/getWeb' });
