param name string
param location string = resourceGroup().location
param tags object = {}

param allowBlobPublicAccess bool = false
param containers array = []
param kind string = 'StorageV2'
param minimumTlsVersion string = 'TLS1_2'
param sku object = { name: 'Standard_LRS' }
param allowedIpAddresses array = []
// param virtualNetworkSubnetId string

var ipRules = [
  for ipAddress in allowedIpAddresses: {
    action: 'Allow'
    value: ipAddress
  }
]

resource storage 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: name
  location: location
  tags: tags
  kind: kind
  sku: sku
  properties: {
    minimumTlsVersion: minimumTlsVersion
    allowBlobPublicAccess: allowBlobPublicAccess
    publicNetworkAccess: empty(allowedIpAddresses) ? 'Disabled': 'Enabled'
    allowSharedKeyAccess: false
    defaultToOAuthAuthentication: true
    networkAcls: {
      bypass: 'AzureServices'
      defaultAction: 'Deny'
      ipRules: empty(allowedIpAddresses) ? [] : ipRules
      // Not necessary since a private endpoint ensures the connection between the apps ervice and this storage account
      // virtualNetworkRules: [
      //   {
      //     id: virtualNetworkSubnetId
      //   }
      // ]
    }
  }

  resource blobServices 'blobServices' = if (!empty(containers)) {
    name: 'default'
    resource container 'containers' = [
      for container in containers: {
        name: container.name
        properties: {
          publicAccess: container.?publicAccess ?? 'None'
        }
      }
    ]
  }
}

output name string = storage.name
output primaryEndpoints object = storage.properties.primaryEndpoints
