param name string
param location string = resourceGroup().location
param tags object = {}
param tenantId string

param sku object = { family: 'A', name: 'standard' }
param allowedIpAddresses array = []
param enableSoftDelete bool = true
param virtualNetworkSubnetId string

var ipRules = [
  for ipAddress in allowedIpAddresses: {
    value: ipAddress
  }
]

resource vault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: name
  location: location
  tags: tags
  properties: {
    sku: sku
    tenantId: tenantId
    enableSoftDelete: enableSoftDelete
    enableRbacAuthorization: true
    publicNetworkAccess: empty(allowedIpAddresses) ? 'Disabled': 'Enabled'
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
}

output id string = vault.id
output name string = vault.name
