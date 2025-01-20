# HTTP requests for the Azure function app

Open this collection with HTTP debugger [Bruno](https://www.usebruno.com/) to call your Azure function app and manage your webhooks.

## Prerequisites

Create file `.env` in this folder, paste the following content, and replace the placeholders with your own values:

```env
funchost=YOUR_AZURE_FUNCTION_APP_NAME
code=APP_KEY_VALUE
notificationUrl=https://YOUR_AZURE_FUNCTION_APP_NAME.azurewebsites.net/api/webhooks/service?code=APP_KEY_VALUE
```
