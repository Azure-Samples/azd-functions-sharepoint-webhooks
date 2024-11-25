import { app } from "@azure/functions";
import { registerWebhook, listWehhooks, wehhookService, removeWehhook, showWehhook } from "../webhook/webhook.js"

app.http('webhook-register', { methods: ['POST'], authLevel: 'function', handler: registerWebhook, route: 'webhook/register' });
app.http('webhook-service', { methods: ['POST'], authLevel: 'function', handler: wehhookService, route: 'webhook/service' });
app.http('webhook-list', { methods: ['GET'], authLevel: 'function', handler: listWehhooks, route: 'webhook/list' });
app.http('webhook-remove', { methods: ['POST'], authLevel: 'function', handler: removeWehhook, route: 'webhook/remove' });
app.http('webhook-show', { methods: ['GET'], authLevel: 'function', handler: showWehhook, route: 'webhook/show' });
