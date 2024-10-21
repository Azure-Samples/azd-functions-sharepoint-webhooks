targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the the environment which is used to generate a short unique hash used in all resources.')
param environmentName string

@minLength(1)
@description('Primary location for all resources')
@allowed(['australiaeast', 'eastasia', 'eastus', 'eastus2', 'northeurope', 'southcentralus', 'southeastasia', 'swedencentral', 'uksouth', 'westus2', 'eastus2euap'])
@metadata({
  azd: {
    type: 'location'
  }
})
param location string

@description('List of the public IP addresses allowed to connect to the storage account and the key vault.')
param allowedIpAddresses array

@description('List of the environment variables to create in the Azure functions service.')
param appSettings object

param resourceGroupName string = ''
param appServiceName string = ''
@allowed(['SystemAssigned', 'UserAssigned'])
param appServiceIdentityType string = 'SystemAssigned'
param appUserAssignedIdentityName string = ''
param appServicePlanName string = ''
param applicationInsightsName string = ''
param logAnalyticsName string = ''
param storageAccountName string = ''
param vNetName string = ''
param vaultName string = ''
param disableLocalAuth bool = true
param publicNetworkAccess string = 'Enabled'
param keyVaultEnableSoftDelete bool = true

var abbrs = loadJsonContent('./abbreviations.json')
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = { 'azd-env-name': environmentName }

// Organize resources in a resource group
resource rg 'Microsoft.Resources/resourceGroups@2021-04-01' = {
  name: !empty(resourceGroupName) ? resourceGroupName : '${abbrs.resourcesResourceGroups}${environmentName}'
  location: location
  tags: tags
}

// User assigned managed identity to be used by the Function App to reach storage and service bus
module processorUserAssignedIdentity './core/identity/userAssignedIdentity.bicep' = if (appServiceIdentityType == 'UserAssigned') {
  name: 'processorUserAssignedIdentity'
  scope: rg
  params: {
    location: location
    tags: tags
    identityName: !empty(appUserAssignedIdentityName) ? appUserAssignedIdentityName : '${abbrs.managedIdentityUserAssignedIdentities}processor-${resourceToken}'
  }
}

// The application backend
module processor './app/processor.bicep' = {
  name: 'processor'
  scope: rg
  params: {
    name: !empty(appServiceName) ? appServiceName : '${abbrs.webSitesFunctions}processor-${resourceToken}'
    location: location
    tags: tags
    applicationInsightsName: monitoring.outputs.applicationInsightsName
    appServicePlanId: appServicePlan.outputs.id
    runtimeName: 'node'
    runtimeVersion: '20'
    storageAccountName: storage.outputs.name
    identityType: appServiceIdentityType
    identityId: appServiceIdentityType == 'UserAssigned' ? processorUserAssignedIdentity.outputs.identityId : ''
    identityClientId: appServiceIdentityType == 'UserAssigned' ? processorUserAssignedIdentity.outputs.identityClientId : ''
    appSettings: appSettings
    virtualNetworkSubnetId: serviceVirtualNetwork.outputs.appSubnetID
  }
}

// Backing storage for Azure functions processor
module storage './core/storage/storage-account.bicep' = {
  name: 'storage'
  scope: rg
  params: {
    name: !empty(storageAccountName) ? storageAccountName : '${abbrs.storageStorageAccounts}${resourceToken}'
    location: location
    tags: tags
    containers: [{name: 'deploymentpackage'}]
    publicNetworkAccess: publicNetworkAccess
    allowedIpAddresses:allowedIpAddresses
    virtualNetworkSubnetId: serviceVirtualNetwork.outputs.appSubnetID
  }
}

var storageRoleDefinitionId  = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b' //Storage Blob Data Owner role

// Allow access from processor to storage account using a managed identity
module storageRoleAssignmentApi 'app/storage-Access.bicep' = {
  name: 'storageRoleAssignmentPRocessor'
  scope: rg
  params: {
    storageAccountName: storage.outputs.name
    roleDefinitionID: storageRoleDefinitionId
    principalID: appServiceIdentityType == 'UserAssigned' ? processorUserAssignedIdentity.outputs.identityPrincipalId : processor.outputs.SERVICE_API_IDENTITY_PRINCIPAL_ID
  }
}

module appServicePlan './core/host/appserviceplan.bicep' = {
  name: 'appserviceplan'
  scope: rg
  params: {
    name: !empty(appServicePlanName) ? appServicePlanName : '${abbrs.webServerFarms}${resourceToken}'
    location: location
    tags: tags
    sku: {
      name: 'FC1'
      tier: 'FlexConsumption'
    }
  }
}

// Virtual Network & private endpoint
module serviceVirtualNetwork 'app/vnet.bicep' = {
  name: 'serviceVirtualNetwork'
  scope: rg
  params: {
    location: location
    tags: tags
    vNetName: !empty(vNetName) ? vNetName : '${abbrs.networkVirtualNetworks}${resourceToken}'
  }
}

module servicePrivateEndpoint 'app/storage-PrivateEndpoint.bicep' = {
  name: 'servicePrivateEndpoint'
  scope: rg
  params: {
    location: location
    tags: tags
    virtualNetworkName: !empty(vNetName) ? vNetName : '${abbrs.networkVirtualNetworks}${resourceToken}'
    subnetName: serviceVirtualNetwork.outputs.peSubnetName
    resourceName: storage.outputs.name
  }
}

// Monitor application with Azure Monitor
module monitoring './core/monitor/monitoring.bicep' = {
  name: 'monitoring'
  scope: rg
  params: {
    location: location
    tags: tags
    logAnalyticsName: !empty(logAnalyticsName) ? logAnalyticsName : '${abbrs.operationalInsightsWorkspaces}${resourceToken}'
    applicationInsightsName: !empty(applicationInsightsName) ? applicationInsightsName : '${abbrs.insightsComponents}${resourceToken}'
    disableLocalAuth: disableLocalAuth
  }
}

var monitoringRoleDefinitionId = '3913510d-42f4-4e42-8a64-420c390055eb' // Monitoring Metrics Publisher role ID

// Allow access from processor to application insights using a managed identity
module appInsightsRoleAssignmentApi './core/monitor/appinsights-access.bicep' = {
  name: 'appInsightsRoleAssignmentPRocessor'
  scope: rg
  params: {
    appInsightsName: monitoring.outputs.applicationInsightsName
    roleDefinitionID: monitoringRoleDefinitionId
    principalID: appServiceIdentityType == 'UserAssigned' ? processorUserAssignedIdentity.outputs.identityPrincipalId : processor.outputs.SERVICE_API_IDENTITY_PRINCIPAL_ID
  }
}

// Keyvault
module vault './core/vault/vault-resource.bicep' = {
  name: 'vault'
  scope: rg
  params: {
    name: !empty(vaultName) ? vaultName : '${abbrs.vaultAccounts}${resourceToken}'
    location: location
    tags: tags
    publicNetworkAccess: publicNetworkAccess
    allowedIpAddresses: allowedIpAddresses
    virtualNetworkSubnetId: serviceVirtualNetwork.outputs.appSubnetID
    tenantId: tenant().tenantId
    enableSoftDelete: keyVaultEnableSoftDelete
  }
}

@description('This is the built-in Key Vault Secrets User role. See https://docs.microsoft.com/azure/role-based-access-control/built-in-roles#key-vault-administrator')
resource keyVaultSecretsUserRoleDefinition 'Microsoft.Authorization/roleDefinitions@2022-05-01-preview' existing = {
  scope: subscription()
  name: '4633458b-17de-408a-b874-0445c86b69e6'
}

// Allow access from processor to skey vault using a managed identity
module vaultRoleAssignmentApi './core/vault/vault-access.bicep' = {
  name: 'vaultRoleAssignmentProcessor'
  scope: rg
  params: {
    keyVaultName: vault.outputs.name
    roleDefinitionID: keyVaultSecretsUserRoleDefinition.id
    principalID: appServiceIdentityType == 'UserAssigned' ? processorUserAssignedIdentity.outputs.identityPrincipalId : processor.outputs.SERVICE_API_IDENTITY_PRINCIPAL_ID
  }
}

// App outputs
output APPLICATIONINSIGHTS_CONNECTION_STRING string = monitoring.outputs.applicationInsightsConnectionString
output AZURE_LOCATION string = location
output AZURE_TENANT_ID string = tenant().tenantId
output SERVICE_PROCESSOR_NAME string = processor.outputs.SERVICE_PROCESSOR_NAME
output AZURE_FUNCTION_NAME string = processor.outputs.SERVICE_PROCESSOR_NAME
