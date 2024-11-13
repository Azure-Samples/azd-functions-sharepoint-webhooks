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

This quickstart uses `DefaultAzureCredential`, so the identity the functions service uses to authenticate to SharePoint depends if it runs on the local environment, or in Azure.  
I strongly recommend to read [this article](https://aka.ms/azsdk/js/identity/credential-chains#use-defaultazurecredential-for-flexibility) to understand the concept, if this is a new topic.

## Grant permission to SharePoint when the functions run on the local environment

`DefaultAzureCredential` will preferentially use the delegated credentials of `Azure CLI` to authenticate to SharePoint.  
The PowerShell script below grants the `Azure CLI`'s service principal the SharePoint delegated permission `AllSites.Manage`, which is the minimum required to register a webhook:

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
> The service principal for `Azure CLI` may not exist in your tenant. If so, [this issue](https://github.com/Azure/azure-cli/issues/28628) will help you to add it.

### Grant a delegated SharePoint API permission to the service principal

Since `Sites.Selected` permission does not exist in this context, we will grant the delegated permission `AllSites.FullControl` to the `Azure CLI` service principal.

## Grant permission to SharePoint when the functions run in Azure

The functions service will use a managed identity to authenticate to SharePoint. This may be the existing, system-assigned managed identity of the functions service, or use your own user-assigned managed identity, that you create and assign to the functions service.  
This tutorial will assume that the system-assigned managed identity is used.

### Grant SharePoint API permission Sites.Selected to the service principal

### Grant effective permission on a SharePoint site to the service principal
