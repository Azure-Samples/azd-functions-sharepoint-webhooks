Open this collection with HTTP debugger [Bruno](https://www.usebruno.com/) to execute the HTTP requests.

## Prerequisites

Create file `.env` in this folder, paste the following content, and replace the placeholders with your own values:

```env
funchost=YOUR_AZURE_FUNCTION_APP_NAME
code=APP_KEY_VALUE
notificationUrl=https://YOUR_AZURE_FUNCTION_APP_NAME.azurewebsites.net/api/webhooks/service?code=APP_KEY_VALUE
```
