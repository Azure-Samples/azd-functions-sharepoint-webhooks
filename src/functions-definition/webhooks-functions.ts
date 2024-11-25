import { app } from "@azure/functions";
import { registerWebhook, listWehhooks, wehhookService, removeWehhook, showWehhook } from "../webhooks/webhooks-app.js"

app.http('webhooks-register', { methods: ['POST'], authLevel: 'function', handler: registerWebhook, route: 'webhooks/register' });
app.http('webhooks-service', { methods: ['POST'], authLevel: 'function', handler: wehhookService, route: 'webhooks/service' });
app.http('webhooks-list', { methods: ['GET'], authLevel: 'function', handler: listWehhooks, route: 'webhooks/list' });
app.http('webhooks-remove', { methods: ['POST'], authLevel: 'function', handler: removeWehhook, route: 'webhooks/remove' });
app.http('webhooks-show', { methods: ['GET'], authLevel: 'function', handler: showWehhook, route: 'webhooks/show' });
