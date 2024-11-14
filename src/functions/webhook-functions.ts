import { app } from "@azure/functions";
import { registerWebhook, listRegisteredWehhooks, wehhookService, removeRegisteredWehhook, showRegisteredWehhook } from "../webhook/webhook.js"

app.http('webhook-register', { methods: ['POST'], authLevel: 'function', handler: registerWebhook, route: 'webhook/register' });
app.http('webhook-service', { methods: ['POST'], authLevel: 'function', handler: wehhookService, route: 'webhook/service' });
app.http('webhook-listRegistered', { methods: ['GET'], authLevel: 'function', handler: listRegisteredWehhooks, route: 'webhook/listRegistered' });
app.http('webhook-removeRegistered', { methods: ['POST'], authLevel: 'function', handler: removeRegisteredWehhook, route: 'webhook/removeRegistered' });
app.http('webhook-showRegistered', { methods: ['GET'], authLevel: 'function', handler: showRegisteredWehhook, route: 'webhook/showRegistered' });
