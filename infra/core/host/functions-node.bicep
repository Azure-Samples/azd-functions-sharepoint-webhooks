param name string
param location string = resourceGroup().location
param tags object = {}

// Reference Properties
param applicationInsightsName string = ''
param appServicePlanId string
param storageAccountName string
param virtualNetworkSubnetId string = ''
@allowed(['SystemAssigned', 'UserAssigned'])
param identityType string
@description('User assigned identity name')
param identityId string
param httpsOnly bool = true
@allowed(['Flex', 'Premium'])
param appFunctionType string = 'Flex'

@allowed(['node'])
param runtimeName string
@allowed(['20'])
param runtimeVersion string
param kind string = 'functionapp,linux'

// Microsoft.Web/sites/config
param appSettings object = {}
param instanceMemoryMB int = 2048
param maximumInstanceCount int = 100

var userAssignedIdentities = identityType == 'UserAssigned'
  ? {
      type: identityType
      userAssignedIdentities: {
        '${identityId}': {}
      }
    }
  : {
      type: identityType
    }

var appSettingsPerFunctionType = appFunctionType == 'Premium'
  ? {
      FUNCTIONS_WORKER_RUNTIME: runtimeName
      FUNCTIONS_EXTENSION_VERSION: '~4'
    }
  : {}

resource stg 'Microsoft.Storage/storageAccounts@2022-09-01' existing = {
  name: storageAccountName
}

resource functions 'Microsoft.Web/sites@2023-12-01' = {
  name: name
  location: location
  tags: tags
  kind: kind
  identity: userAssignedIdentities
  properties: {
    serverFarmId: appServicePlanId
    httpsOnly: httpsOnly
    functionAppConfig: appFunctionType == 'Premium'
      ? null
      : {
          deployment: {
            storage: {
              type: 'blobContainer'
              value: '${stg.properties.primaryEndpoints.blob}deploymentpackage'
              authentication: {
                type: identityType == 'SystemAssigned' ? 'SystemAssignedIdentity' : 'UserAssignedIdentity'
                userAssignedIdentityResourceId: identityType == 'UserAssigned' ? identityId : ''
              }
            }
          }
          scaleAndConcurrency: {
            instanceMemoryMB: instanceMemoryMB
            maximumInstanceCount: maximumInstanceCount
          }
          runtime: {
            name: runtimeName
            version: runtimeVersion
          }
        }
    virtualNetworkSubnetId: !empty(virtualNetworkSubnetId) ? virtualNetworkSubnetId : null
    keyVaultReferenceIdentity: identityType == 'UserAssigned' ? identityId : 'SystemAssigned'

    // Base on both links below, properties vnetRouteAllEnabled anbd vnetContentShareEnabled are not needed for flex functions:
    // https://learn.microsoft.com/en-us/azure/azure-functions/functions-app-settings#flex-consumption-plan-deprecations and https://learn.microsoft.com/en-us/azure/azure-functions/functions-networking-options?tabs=azure-portal#outbound-ip-restrictions
    // But still required for other functions as a workaround to access network-restricted vaults: https://learn.microsoft.com/en-us/azure/app-service/app-service-key-vault-references?tabs=azure-cli#access-network-restricted-vaults
    vnetRouteAllEnabled: true
    vnetContentShareEnabled: true

    siteConfig: {
      keyVaultReferenceIdentity: identityType == 'UserAssigned' ? identityId : 'SystemAssigned'
      linuxFxVersion: appFunctionType == 'Premium' ? '${runtimeName}|${runtimeVersion}' : null
      vnetRouteAllEnabled: true // see above
    }
  }

  resource configAppSettings 'config' = {
    name: 'appsettings'
    properties: union(
      appSettings,
      {
        AzureWebJobsStorage__accountName: stg.name
        AzureWebJobsStorage__credential: 'managedidentity'
        APPLICATIONINSIGHTS_CONNECTION_STRING: applicationInsights.properties.ConnectionString
      },
      appSettingsPerFunctionType
    )
  }
}

resource applicationInsights 'Microsoft.Insights/components@2020-02-02' existing = if (!empty(applicationInsightsName)) {
  name: applicationInsightsName
}

output name string = functions.name
output uri string = 'https://${functions.properties.defaultHostName}'
output identityPrincipalId string = identityType == 'SystemAssigned' ? functions.identity.principalId : ''
