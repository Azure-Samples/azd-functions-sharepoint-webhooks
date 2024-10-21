param name string
param location string = resourceGroup().location
param tags object = {}
param tenantId string

@allowed(['Enabled', 'Disabled'])
param publicNetworkAccess string = 'Disabled'
param sku object = { family: 'A', name: 'standard' }
param allowedIpAddresses array
param virtualNetworkSubnetId string
param enableSoftDelete bool = true

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
    publicNetworkAccess: publicNetworkAccess
    networkAcls: {
      bypass: 'AzureServices'
      defaultAction: 'Deny'
      ipRules: ipRules
      // virtualNetworkRules: map([virtualNetworkSubnetId], subnetId => { id: subnetId })
      virtualNetworkRules: [
        {
          id: virtualNetworkSubnetId
        }
      ]
    }
  }
}

output id string = vault.id
output name string = vault.name
