param name string
param location string = resourceGroup().location
param tags object = {}
param applicationInsightsName string = ''
param appServicePlanId string
param appSettings object = {}
param runtimeName string
param runtimeVersion string
param serviceName string = 'processor'
param storageAccountName string
param virtualNetworkSubnetId string = ''
param instanceMemoryMB int = 2048
param maximumInstanceCount int = 100
param identityId string = ''
param identityClientId string = ''
@allowed(['SystemAssigned', 'UserAssigned'])
param identityType string

var managedIdentityAuthSettings = identityType == 'UserAssigned'
  ? {
      AzureWebJobsStorage__clientId: identityClientId
      APPLICATIONINSIGHTS_AUTHENTICATION_STRING: 'ClientId=${identityClientId};Authorization=AAD'
    }
  : {
      APPLICATIONINSIGHTS_AUTHENTICATION_STRING: 'Authorization=AAD'
    }

module processor '../core/host/functions-flexconsumption.bicep' = {
  name: '${serviceName}-functions-module'
  params: {
    name: name
    location: location
    tags: union(tags, { 'azd-service-name': serviceName })
    identityType: identityType
    identityId: identityId
    appSettings: union(appSettings, managedIdentityAuthSettings)
    applicationInsightsName: applicationInsightsName
    appServicePlanId: appServicePlanId
    runtimeName: runtimeName
    runtimeVersion: runtimeVersion
    storageAccountName: storageAccountName
    virtualNetworkSubnetId: virtualNetworkSubnetId
    instanceMemoryMB: instanceMemoryMB
    maximumInstanceCount: maximumInstanceCount
  }
}

output SERVICE_PROCESSOR_NAME string = processor.outputs.name
output SERVICE_API_IDENTITY_PRINCIPAL_ID string = processor.outputs.identityPrincipalId
