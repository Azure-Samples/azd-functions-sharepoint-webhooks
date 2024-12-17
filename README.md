---
name: Azure Functions for SharePoint Online
description: This quickstart uses azd CLI to deploy Azure Functions which can connect to your own SharePoint Online tenant.
page_type: sample
languages:
- azdeveloper
- bicep
- nodejs
- typescript
products:
- azure-functions
- sharepoint-online
urlFragment: functions-quickstart-spo-azd
---

# Azure Functions for SharePoint Online

This quickstart is based on [this repository](https://github.com/Azure-Samples/functions-quickstart-typescript-azd). It uses Azure Developer command-line (azd) tools to deploy Azure Functions which can list, register and process [SharePoint Online webhooks](https://learn.microsoft.com/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks) on your own tenant.  
The Azure functions use the [Flex Consumption plan](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-plan), are written in TypeScript and run in Node.js 20.  
The popular library [PnPjs](https://pnp.github.io/pnpjs/) is used to interact with SharePoint.  

## Overview

5 HTTP-triggered functions are deployed to show, list, register, process and remove webhooks.  
When receiving a notification from SharePoint, the service function will add a new item to the list `webhookHistory` (can be changed in environment variable `WebhookHistoryListTitle`). It will also record the event in Application Insights.

## Security of the Azure resources

The resources deployed in Azure are configured with a high level of security: 
- The functions service connects to the storage account and the key vault using a private endpoint.
- No network access is allowed on the storage account and the key vault, except on specified IPs (configurable).
- Authorization is configured using the functions service's managed identity (no access key or legacy access policy is enabled).
- All the functions require a key to be called.

## Prerequisites

+ [Node.js 20](https://www.nodejs.org/)
+ [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local?pivots=programming-language-typescript#install-the-azure-functions-core-tools)
+ [Azure Developer CLI (AZD)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
+ Have the built-in role `Owner`, or `Contributor` + [`Role Based Access Control Administrator`](https://learn.microsoft.com/en-us/azure/role-based-access-control/built-in-roles/privileged#role-based-access-control-administrator), to successfully assign roles to the managed identity, as part of the provisioning process
+ To use Visual Studio Code to run and debug locally:
  + [Visual Studio Code](https://code.visualstudio.com/)
  + [Azure Functions extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions)

## Initialize the local project

You can initialize a project from this `azd` template in one of these ways:

+ Use this `azd init` command from an empty local (root) folder:

    ```shell
    azd init --template Yvand/functions-quickstart-spo-azd
    ```

    Supply an environment name, such as `spofuncs-quickstart` when prompted. In `azd`, the environment is used to maintain a unique deployment context for your app.

+ Clone the GitHub template repository, and create an `azd` environment (in this example, `spofuncs-quickstart`):

    ```shell
    git clone https://github.com/Yvand/functions-quickstart-spo-azd.git
    cd functions-quickstart-spo-azd
    azd env new spofuncs-quickstart
    ```

## Prepare your local environment

1. Add a file named `local.settings.json` in the root of your project with the following contents:

   ```json
   {
      "IsEncrypted": false,
      "Values": {
         "AzureWebJobsStorage": "UseDevelopmentStorage=true",
         "FUNCTIONS_WORKER_RUNTIME": "node",
         "TenantPrefix": "YOUR_SHAREPOINT_TENANT_PREFIX",
         "SiteRelativePath": "/sites/YOUR_SHAREPOINT_SITE_NAME"
      }
   }
   ```

1. Review the file `infra/main.parameters.json` to customize the parameters used for provisioning the resources in Azure. Review [this article](https://learn.microsoft.com/en-us/azure/developer/azure-developer-cli/manage-environment-variables) to manage the azd's environment variables.

   Important: Ensure the values for `TenantPrefix` and `SiteRelativePath` are identical between the files `local.settings.json` (used when running the functions locally) and `infra\main.parameters.json` (used to set the environment variables in Azure).

1. Install the dependencies and build the functions app:

   ```shell
   npm install
   npm run build
   ```

1. Provision the resources in Azure and deploy the functions app package by running command `azd up`.

1. The functions can also be run locally by executing command `npm run start`.

# Grant the functions access to SharePoint Online

The authentication to SharePoint is done using `DefaultAzureCredential`, so the credential used depends if the functions run on your local environment, or in Azure.  
If you never heard about `DefaultAzureCredential`, you should familirize yourself with its concept by reading [this article](https://aka.ms/azsdk/js/identity/credential-chains#use-defaultazurecredential-for-flexibility), before continuing.

## Grant the functions access to SharePoint when they run on the local environment

`DefaultAzureCredential` will preferentially use the delegated credentials of `Azure CLI` to authenticate to SharePoint.  
Use the Microsoft Graph PowerShell script below to grant the SharePoint delegated permission `AllSites.Manage` to the `Azure CLI`'s service principal:

```powershell
Connect-MgGraph -Scope "Application.Read.All", "DelegatedPermissionGrant.ReadWrite.All"
$scopeName = "AllSites.Manage"
$requestorAppPrincipalObj = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Azure CLI'"
$resourceAppPrincipalObj = Get-MgServicePrincipal -Filter "displayName eq 'Office 365 SharePoint Online'"

$params = @{
  clientId = $requestorAppPrincipalObj.Id
  consentType = "AllPrincipals"
  resourceId = $resourceAppPrincipalObj.Id
  scope = $scopeName
}
New-MgOauth2PermissionGrant -BodyParameter $params
```

> [!WARNING]  
> The service principal for `Azure CLI` may not exist in your tenant. If so, check [this issue](https://github.com/Azure/azure-cli/issues/28628) to add it.

> [!NOTE]  
> `AllSites.Manage` is the minimum permission required to register a webhook.
> `Sites.Selected` cannot be used because it does not exist as a delegated permission in the SharePoint API.

## Grant the functions access to SharePoint when they run in Azure

`DefaultAzureCredential` will use a managed identity to authenticate to SharePoint. This may be the existing, system-assigned managed identity of the functions service, or a user-assigned managed identity.  
This tutorial will assume that the system-assigned managed identity is used.

### Grant the SharePoint API permission Sites.Selected to the managed identity

Navigate to the [function apps in the Azure portal](https://portal.azure.com/#blade/HubsExtension/BrowseResourceBlade/resourceType/Microsoft.Web%2Fsites/kind/functionapp) > Select your app > Identity. Note the `Object (principal) ID` of the system-assigned managed identity.  
In this tutorial, it is `d3e8dc41-94f2-4b0f-82ff-ed03c363f0f8`.  
Then, use one of the scripts below to grant it the app-only permission `Sites.Selected` on the SharePoint API:

<details>
  <summary>Using the Microsoft Graph PowerShell SDK</summary>

```powershell
# This script requires the modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, which can be installed with the cmdlet Install-Module below:
# Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Repository PSGallery -Force
Connect-MgGraph -Scope "Application.Read.All", "AppRoleAssignment.ReadWrite.All"
$managedIdentityObjectId = "d3e8dc41-94f2-4b0f-82ff-ed03c363f0f8" # 'Object (principal) ID' of the managed identity
$scopeName = "Sites.Selected"
$resourceAppPrincipalObj = Get-MgServicePrincipal -Filter "displayName eq 'Office 365 SharePoint Online'" # SPO
$targetAppPrincipalAppRole = $resourceAppPrincipalObj.AppRoles | ? Value -eq $scopeName

$appRoleAssignment = @{
    "principalId" = $managedIdentityObjectId
    "resourceId"  = $resourceAppPrincipalObj.Id
    "appRoleId"   = $targetAppPrincipalAppRole.Id
}
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentityObjectId -BodyParameter $appRoleAssignment | Format-List
```

</details>
   
<details>
  <summary>Using az cli in Bash</summary>

```bash
managedIdentityObjectId="d3e8dc41-94f2-4b0f-82ff-ed03c363f0f8" # 'Object (principal) ID' of the managed identity
resourceServicePrincipalId=$(az ad sp list --query '[].[id]' --filter "displayName eq 'Office 365 SharePoint Online'" -o tsv)
resourceServicePrincipalAppRoleId="$(az ad sp show --id $resourceServicePrincipalId --query "appRoles[?starts_with(value, 'Sites.Selected')].[id]" -o tsv)"

az rest --method POST --uri "https://graph.microsoft.com/v1.0/servicePrincipals/${managedIdentityObjectId}/appRoleAssignments" --headers 'Content-Type=application/json' --body "{ 'principalId': '${managedIdentityObjectId}', 'resourceId': '${resourceServicePrincipalId}', 'appRoleId': '${resourceServicePrincipalAppRoleId}' }"
```

</details>

### Grant the managed identity effective access to a SharePoint site

Navigate to the [Enterprise applications in the Entra ID portal](https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/) > Set the filter `Application type` to `Managed Identities` > Click on your managed identity and note its `Application ID`.  
In this tutorial, it is `3150363e-afbe-421f-9785-9d5404c5ae34`.  

> [!WARNING]  
> In this step, we will use the `Application ID` of the managed identity, while in the previous step we used its `Object ID`, be mindful about the risk of confusion.

Then, use one of the scripts below to grant it the app-only permission `manage` on a specific SharePoint site:

> [!NOTE]  
> The managed identity of the functions service is granted SharePoint permission `manage`, because it is the minimum required to register a webhook.

<details>
  <summary>Using PnP PowerShell</summary>

[PnP PowerShell](https://pnp.github.io/powershell/cmdlets/Grant-PnPAzureADAppSitePermission.html)

```powershell
Connect-PnPOnline -Url "https://YOUR_SHAREPOINT_TENANT_PREFIX.sharepoint.com/sites/YOUR_SHAREPOINT_SITE_NAME" -Interactive -ClientId "YOUR_PNP_APP_CLIENT_ID"
Grant-PnPAzureADAppSitePermission -AppId "3150363e-afbe-421f-9785-9d5404c5ae34" -DisplayName "YOUR_FUNC_APP_NAME" -Permissions Manage
```

</details>
   
<details>
  <summary>Using m365 cli in Bash</summary>

[m365 cli](https://pnp.github.io/cli-microsoft365/cmd/spo/site/site-apppermission-add/)

```bash
targetapp="3150363e-afbe-421f-9785-9d5404c5ae34"
siteUrl="https://YOUR_SHAREPOINT_TENANT_PREFIX.sharepoint.com/sites/YOUR_SHAREPOINT_SITE_NAME"
m365 spo site apppermission add --appId $targetapp --permission manage --siteUrl $siteUrl
```

</details>

> [!IMPORTANT]  
> The app registration used to run those commands must have at least the following permissions:
> - Delegated permission `Application.ReadWrite.All` in the Graph API
> - Delegated permission `AllSites.FullControl` in the SharePoint API

## Call the functions

For security reasons, when running in Azure, functions require an app key to pass in query string parameter `code`. The app keys can be found in the functions app service > App Keys.  
Most functions take optional parameters `tenantPrefix` and `siteRelativePath`. If they are not specified, the values set in the app's environment variables will be used.

### Using vscode extension RestClient

You can use the Visual Studio Code extension [`REST Client`](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) to execute the requests in the .http file.  
It takes parameters from a .env file on the same folder. You can create it based on the sample files `azure.env.example` and `local.env.example`.

### Using curl

Below is a sample script in Bash that calls the functions in Azure using `curl`:

```bash
# Edit those variables to fit your app function
funchost="YOUR_FUNC_APP_NAME"
code="YOUR_HOST_KEY"
notificationUrl="https://${funchost}.azurewebsites.net/api/webhooks/service?code=${code}"
listTitle="YOUR_SHAREPOINT_LIST"

# List all the webhooks registered on a list
curl "https://${funchost}.azurewebsites.net/api/webhooks/list?code=${code}&listTitle=${listTitle}"

# Register a webhook
curl -X POST "https://${funchost}.azurewebsites.net/api/webhooks/register?code=${code}&listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Show this webhook registered on a list
curl "https://${funchost}.azurewebsites.net/api/webhooks/show?code=${code}&listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Remove the webhook from the list
# Step 1: Get the webhook id in the output of the function /webhooks/show
webhookId=$(curl -s "https://${funchost}.azurewebsites.net/api/webhooks/show?code=${code}&listTitle=${listTitle}&notificationUrl=${notificationUrl}" | \
    python3 -c "import sys, json; document = json.load(sys.stdin); document and print(document['id'])")
# Step 2: Call function /webhooks/remove and pass the webhookId
curl -X POST "https://${funchost}.azurewebsites.net/api/webhooks/remove?code=${code}&listTitle=${listTitle}&webhookId=${webhookId}"
```

The same script, which calls the functions when they run in your local environment:

```bash
# Edit those variables to fit your app function
funchost="YOUR_FUNC_APP_NAME"
code="YOUR_HOST_KEY"
notificationUrl="https://${funchost}.azurewebsites.net/api/webhooks/service?code=${code}"
listTitle="YOUR_SHAREPOINT_LIST"

# List all the webhooks registered on a list
curl "http://localhost:7071/api/webhooks/list?listTitle=${listTitle}"

# Register a webhook
curl -X POST "http://localhost:7071/api/webhooks/register?listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Show this webhook registered on a list
curl "http://localhost:7071/api/webhooks/show?listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Remove the webhook from the list
# Step 1: Get the webhook id in the output of the function /webhooks/show
webhookId=$(curl -s "http://localhost:7071/api/webhooks/show?listTitle=${listTitle}&notificationUrl=${notificationUrl}" | \
    python3 -c "import sys, json; document = json.load(sys.stdin); document and print(document['id'])")
# Step 2: Call function /webhooks/remove and pass the webhookId
curl -X POST "http://localhost:7071/api/webhooks/remove?listTitle=${listTitle}&webhookId=${webhookId}"
```

## Review the logs

When the functions run in your local environment, the logging goes to the console.  
When the functions run in Azure, the logging goes to the Application Insights resource configured in the app service.  

### KQL queries for Application Insights

The KQL query below shows the entries from all the functions, and filters out the logging from the infrastructure:

```kql
traces 
| where isnotempty(operation_Name)
| project timestamp, operation_Name, severityLevel, message
| order by timestamp desc
```

The KQL query below does the following:

- Includes only the entries from the function `webhooks/service` (which receives the notifications from SharePoint)
- Parses the `message` as a json document (which is how this project writes the messages)
- Includes only the entries that were successfully parsed (excludes those from the infrastructure)

```kql
traces 
| where operation_Name contains "webhooks-service"
| extend jsonMessage = parse_json(message)
| where isnotempty(jsonMessage.['message'])
| project timestamp, operation_Name, severityLevel, jsonMessage.['message'], jsonMessage.['error']
| order by timestamp desc
```

## Known issues

Azure Functions Flex Consumption plan is currently in preview, be aware about its [current limitations and issues](https://learn.microsoft.com/en-us/azure/azure-functions/flex-consumption-plan#considerations).

## Cleanup the resources in Azure

You can delete all the resources this project created in Azure, by running the command `azd down`.  
Alternatively, you can delete the resource group, which has the azd environment's name by default.
