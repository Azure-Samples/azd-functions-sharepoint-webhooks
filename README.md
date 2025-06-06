<!--
---
name: Azure function app for SharePoint webhooks
description: This quickstart uses azd CLI to deploy an Azure function app which
  connects to your SharePoint Online tenant, to register and manage webhooks, 
  and process the notifications from SharePoint.
page_type: sample
languages:
  - azdeveloper
  - bicep
  - nodejs
  - typescript
products:
  - azure-functions
  - office-sp
urlFragment: azd-functions-sharepoint-webhooks
---
-->

# Azure function app for SharePoint webhooks

This template uses [Azure Developer CLI (azd)](https://aka.ms/azd) to deploy an Azure function app that connects to your SharePoint Online tenant, to register and manage [webhooks](https://learn.microsoft.com/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks), and process the notifications from SharePoint.

## Overview

The function app uses the [Flex Consumption plan](https://learn.microsoft.com/azure/azure-functions/flex-consumption-plan), hosts multiple HTTP-triggered functions written in TypeScript, and uses [PnPjs](https://pnp.github.io/pnpjs/) to communicate with SharePoint.  
When receiving a notification from SharePoint, the function gets all the changes for the past 15 minutes on the list that triggered it, and adds an item to the list **webhookHistory** (created if it does not exist).

## Security of the Azure resources

The resources are deployed in Azure with a high level of security:

- The function app connects to the storage account using a private endpoint.
- No public network access is allowed on the storage account.
- All the permissions are granted to the function app's managed identity (no secret, access key or legacy access policy is used).
- All the functions require an app key to be called.

## Prerequisites

+ [Node.js 22](https://www.nodejs.org/)
- [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
- An Azure subscription that trusts the same Microsoft Entra ID directory as the SharePoint tenant

## Permissions required to provision the resources in Azure

The account running **azd** must have at least the following roles to successfully provision the resources:

- Azure role **[Contributor](https://learn.microsoft.com/azure/role-based-access-control/built-in-roles/privileged#contributor)**: To create all the resources needed
- Azure role **[Role Based Access Control Administrator](https://learn.microsoft.com/azure/role-based-access-control/built-in-roles/privileged#role-based-access-control-administrator)**: To assign roles (to access the storage account and Application Insights) to the managed identity of the function app

## Initialize the project

1. Run **azd init** from an empty local (root) folder:

    ```console
    azd init --template azd-functions-sharepoint-webhooks
    ```

    Enter an environment name, such as **spofuncs-quickstart** when prompted. In **azd**, the environment is used to maintain a unique deployment context for your app.


1. In the root of your project, add a file named **local.settings.json** with the content below, and set the variables `TenantPrefix` and `SiteRelativePath` to match your SharePoint tenant:

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

1. Review the file **infra/main.parameters.json**, and set the variables `TenantPrefix` and `SiteRelativePath` to match your SharePoint tenant.

   Review the article on [Manage environment variables](https://learn.microsoft.com/azure/developer/azure-developer-cli/manage-environment-variables) to manage the azd's environment variables.

1. Install the dependencies and build the function app:

   ```console
   npm install
   npm run build
   ```

## Run the function app

It can run either locally or in Azure:

- To run the function app locally: Run **npm run start**.
- To provision the resources in Azure and deploy the function app: Run **azd up**.

## Grant the function app access to SharePoint Online

The authentication to SharePoint is done using `DefaultAzureCredential`, so the credential used depends on whether the function app runs locally, or in Azure.  

If you never heard about `DefaultAzureCredential`, you should familiarize yourself with its concept by referring to the section **Use DefaultAzureCredential for flexibility** in [Credential chains in the Azure Identity client library for JavaScript](https://learn.microsoft.com/azure/developer/javascript/sdk/authentication/credential-chains).

### When it runs on your local environment

`DefaultAzureCredential` will preferentially use the delegated credentials of **Azure CLI** to authenticate to SharePoint.  

Use the [Microsoft Graph PowerShell](https://learn.microsoft.com/powershell/microsoftgraph/) script below to grant the SharePoint delegated permission **AllSites.Manage** to the **Azure CLI**'s service principal:

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
> - The service principal for **Azure CLI** may not exist in your tenant. If so, check [this issue](https://github.com/Azure/azure-cli/issues/28628) to add it.
> - The scope [`DelegatedPermissionGrant.ReadWrite.All`](https://learn.microsoft.com/graph/permissions-reference#approleassignmentreadwriteall) is necessary to run the script, and requires the admin consent.

> [!NOTE]  
> **AllSites.Manage** is the minimum permission required to register a webhook.
> **Sites.Selected** cannot be used because it does not exist as a delegated permission in the SharePoint API.

### When it runs in Azure

`DefaultAzureCredential` will use a managed identity to authenticate to SharePoint. This may be the existing, system-assigned managed identity of the function app service or a user-assigned managed identity.  

This tutorial assumes the system-assigned managed identity is used.

#### Grant the SharePoint API permission Sites.Selected to the managed identity

Navigate to your function app in the [Azure portal](https://portal.azure.com/#blade/HubsExtension/BrowseResourceBlade/resourceType/Microsoft.Web%2Fsites/kind/functionapp) > select **Identity** and note the **Object (principal) ID** of the system-assigned managed identity.  

> [!NOTE]
> In this tutorial, it is **d3e8dc41-94f2-4b0f-82ff-ed03c363f0f8**.  

Then, use one of the scripts below to grant this identity the app-only permission **Sites.Selected** on the SharePoint API:

> [!IMPORTANT]
> The scripts below require at least the delegated permission [`AppRoleAssignment.ReadWrite.All`](https://learn.microsoft.com/graph/permissions-reference#approleassignmentreadwriteall) (requires admin consent)

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

#### Grant the managed identity effective access to a SharePoint site

Navigate to the [Enterprise applications](https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/) > Set the **Application type** filter to **Managed Identities** > select your managed identity and note its **Application ID**.

> [!NOTE]
> In this tutorial, it is **3150363e-afbe-421f-9785-9d5404c5ae34**.  

Then, use one of the scripts below to grant it the app-only permission **manage** (minimum required to register a webhook) on a specific SharePoint site:

> [!IMPORTANT]  
> The app registration used to run those scripts must have at least the following permissions:
>
> - Delegated permission **Application.ReadWrite.All** in the Graph API (requires admin consent)
> - Delegated permission **AllSites.FullControl** in the SharePoint API (requires admin consent)

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

## Call the function app

For security reasons, when running in Azure, the function app requires an app key to pass in the query string parameter **code**. The app keys are found in the function app service's **App Keys** keys page.  

Most HTTP functions take optional parameters `tenantPrefix` and `siteRelativePath`. If they are not specified, the values in the app's environment variables are used.  

<details>
  <summary>Using API debugger Bruno</summary>

Review [this README](http-requests-collection/README.md) for more information.

</details>

<details>
  <summary>Using vscode extension RestClient</summary>

You can use the Visual Studio Code extension [`REST Client`](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) to execute the requests in the .http file.  
It takes parameters from a .env file on the same folder. You can create it based on the sample files **azure.env.example** and **local.env.example**.

</details>

<details>
  <summary>Using PowerShell</summary>

Below is a sample script in PowerShell that calls the function app using `Invoke-RestMethod`:

```powershell
# Format of the values if calling the function app in Azure
$funchost = "https://<YOUR_FUNC_APP_NAME>.azurewebsites.net"
$code = "code=<YOUR_HOST_KEY>&"

# Format of the values if calling the function app locally (for debugging)
$funchost = "http://localhost:7071"
$code = ""

# Other variables
$listTitle = "<YOUR_SHAREPOINT_LIST_NAME>"
$notificationUrl = "https://<YOUR_FUNC_APP_NAME>.azurewebsites.net/api/webhooks/service?code=<YOUR_HOST_KEY>"

# List all the webhooks registered on a list
Invoke-RestMethod -Method GET -Uri "${funchost}/api/webhooks/list?${code}listTitle=${listTitle}"

# Register a webhook in a list
Invoke-RestMethod -Method POST -Uri "${funchost}/api/webhooks/register?${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Show this webhook registered on a list
Invoke-RestMethod -Method GET -Uri "${funchost}/api/webhooks/show?${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Remove the webhook from the list
# Step 1: Call the function /webhooks/show to get the webhook id
$webhookId = $(Invoke-RestMethod -Method GET -Uri "${funchost}/api/webhooks/show?${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}").Id
# Step 2: Call the function /webhooks/remove and pass the webhook id
Invoke-RestMethod -Method POST -Uri "${funchost}/api/webhooks/remove?${code}listTitle=${listTitle}&webhookId=${webhookId}"
```

</details>

<details>
  <summary>Using curl</summary>

Below is a sample script in Bash that calls the function app using `curl`:

```bash
# Format of the values if calling the function app in Azure
funchost="https://<YOUR_FUNC_APP_NAME>.azurewebsites.net"
code="code=<YOUR_HOST_KEY>&"

# Format of the values if calling the function app locally (for debugging)
funchost="http://localhost:7071"
code=""

# Other variables
listTitle="<YOUR_SHAREPOINT_LIST_NAME>"
notificationUrl="https://<YOUR_FUNC_APP_NAME>.azurewebsites.net/api/webhooks/service?code=<YOUR_HOST_KEY>"

# List all the webhooks registered on a list
curl "${funchost}/api/webhooks/list?${code}listTitle=${listTitle}"

# Register a webhook in a list
curl -X POST "${funchost}/api/webhooks/register?${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Show this webhook registered on a list
curl "${funchost}/api/webhooks/show?code=${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}"

# Remove the webhook from the list
# Step 1: Call the function /webhooks/show to get the webhook id
webhookId=$(curl -s "${funchost}/api/webhooks/show?code=${code}listTitle=${listTitle}&notificationUrl=${notificationUrl}" | \
    python3 -c "import sys, json; document = json.load(sys.stdin); document and print(document['id'])")
# Step 2: Call the function /webhooks/remove and pass the webhook id
curl -X POST "${funchost}/api/webhooks/remove?${code}listTitle=${listTitle}&webhookId=${webhookId}"
```

</details>

## Review the logs

When the function app runs in your local environment, the logging goes to the console.  
When the function app runs in Azure, the logging goes to the Application Insights resource configured in the app service.  

### KQL queries for Application Insights

The KQL query below shows the entries from all the HTTP functions, and filters out the logging from the infrastructure:

```kql
traces 
| where isnotempty(operation_Name)
| project timestamp, operation_Name, severityLevel, message
| order by timestamp desc
```

The KQL query below does the following:

- Includes only the entries from the function **webhooks/service** (which receives the notifications from SharePoint)
- Parses the **message** as a json document (which is how this project writes the messages)
- Includes only the entries that were successfully parsed (that excludes logs from the infrastructure)

```kql
traces 
| where operation_Name contains "webhooks-service"
| extend jsonMessage = parse_json(message)
| where isnotempty(jsonMessage.['message'])
| project timestamp, operation_Name, severityLevel, jsonMessage.['message'], jsonMessage.['error']
| order by timestamp desc
```

## Known issues

The Flex Consumption plan has [limitations](https://learn.microsoft.com/azure/azure-functions/flex-consumption-plan#considerations) you should be aware of.

## Cleanup the resources in Azure

You can delete all the resources this project created in Azure, by running the command **azd down**.  

Alternatively, you can delete the resource group, that has the azd environment's name by default.
