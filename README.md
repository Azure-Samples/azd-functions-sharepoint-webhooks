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

# Secured Azure Functions for SharePoint Online

This quickstart uses Azure Developer command-line (azd) tools to deploy Azure Functions which can list, register and process SharePoint Online webhooks on your own tenant. It uses a managed identity and a virtual network to make sure the deployment is secure by default.

## Prerequisites

+ [Node.js 20](https://www.nodejs.org/)
+ [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local?pivots=programming-language-typescript#install-the-azure-functions-core-tools)
+ [Azure Developer CLI (AZD)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
+ To use Visual Studio Code to run and debug locally:
  + [Visual Studio Code](https://code.visualstudio.com/)
  + [Azure Functions extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions)

## Initialize the local project

You can initialize a project from this `azd` template in one of these ways:

+ Use this `azd init` command from an empty local (root) folder:


    ```shell
    azd init --template Yvand/functions-quickstart-spo-azd
    ```

    Supply an environment name, such as `flexquickstart` when prompted. In `azd`, the environment is used to maintain a unique deployment context for your app.

+ Clone the GitHub template repository locally using the `git clone` command:

    ```shell
    git clone https://github.com/Yvand/functions-quickstart-spo-azd.git
    cd functions-quickstart-spo-azd
    ```

    You can also clone the repository from your own fork in GitHub.

## Prepare your local environment

Add a file named `local.settings.json` in the root of your project with the following contents:

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

# Grant the functions access to SharePoint Online

The authentication to SharePoint is done using `DefaultAzureCredential`, so the credential used depends if the functions run on the local environment, or in Azure.  
If you never heard about `DefaultAzureCredential`, you should familirize yourself with its concept by reading [this article](https://aka.ms/azsdk/js/identity/credential-chains#use-defaultazurecredential-for-flexibility), before continuing.

## Grant permission to SharePoint when the functions run on the local environment

`DefaultAzureCredential` will preferentially use the delegated credentials of `Azure CLI` to authenticate to SharePoint.  
The PowerShell script below grants the SharePoint delegated permission `AllSites.Manage` to the `Azure CLI`'s service principal:

```powershell
Connect-MgGraph -Scope "Application.Read.All", "DelegatedPermissionGrant.ReadWrite.All"
$scopeName = "AllSites.Manage"
$requestorAppPrincipalObj = Get-MgServicePrincipal -Filter "displayname eq 'Microsoft Azure CLI'"
$resourceAppPrincipalObj = Get-MgServicePrincipal -Filter "displayname eq 'Office 365 SharePoint Online'"

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

> [!IMPORTANT]  
> `AllSites.Manage` is the minimum permission required to register a webhook.. `Sites.Selected` cannot be used because it does not exist as a delegated permission in the SharePoint API.

## Grant permission to SharePoint when the functions run in Azure

The functions service will use a managed identity to authenticate to SharePoint. This may be the existing, system-assigned managed identity of the functions service, or use your own user-assigned managed identity, that you create and assign to the functions service.  
This tutorial will assume that the system-assigned managed identity is used.

### Grant SharePoint API permission Sites.Selected to the service principal

TODO

<details>
  <summary>Using PowerShell</summary>
  ```powershell
  # TODO
  ```
</details>
   
<details>
  <summary>Using az cli in Bash</summary>
  ```bash
  managedIdentityObjectId="0efdba91-0b79-461a-af50-377740abf811" # 'Object (principal) ID' of the managed identity
  resourceServicePrincipalId=$(az ad sp list --query '[].[id]' --filter "displayName eq 'Office 365 SharePoint Online'" -o tsv)
  resourceServicePrincipalAppRoleId="$(az ad sp show --id $resourceServicePrincipalId --query "appRoles[?starts_with(value, 'Sites.Selected')].[id]" -o tsv)"
    
  az rest --method POST --uri "https://graph.microsoft.com/v1.0/servicePrincipals/${managedIdentityObjectId}/appRoleAssignments" --headers 'Content-Type=application/json' --body "{ 'principalId': '${managedIdentityObjectId}', 'resourceId': '${resourceServicePrincipalId}', 'appRoleId': '${resourceServicePrincipalAppRoleId}' }"
  ```
</details>

### Grant effective permission on a SharePoint site to the service principal
