// Parameters
@description('Specifies the name of the virtual network.')
param virtualNetworkName string

@description('Specifies the name of the subnet which contains the virtual machine.')
param subnetName string

@description('Specifies the resource name of the vault resource with an endpoint.')
param resourceName string

@description('Specifies the location.')
param location string = resourceGroup().location

param tags object = {}

// Virtual Network
resource vnet 'Microsoft.Network/virtualNetworks@2021-08-01' existing = {
  name: virtualNetworkName
}

resource keyVaultResource 'Microsoft.KeyVault/vaults@2023-07-01' existing = {
  name: resourceName
}

// cannot use format('privatelink{0}', environment().suffixes.keyvaultDns) because keyvaultDns returns 'vault.azure.net' instead of 'vaultcore.azure.net' - https://github.com/Azure/bicep/issues/9839
var vaultPrivateDNSZoneName = 'privatelink.vaultcore.azure.net'

// AVM module for vault Private DNS Zone
module privateDnsZoneVaultDeployment 'br/public:avm/res/network/private-dns-zone:0.7.1' = {
  name: 'vault-private-dns-zone-deployment'
  params: {
    name: vaultPrivateDNSZoneName
    location: 'global'
    tags: tags
    virtualNetworkLinks: [
      {
        name: '${resourceName}-vault-link-${take(toLower(uniqueString(resourceName, virtualNetworkName)), 4)}'
        virtualNetworkResourceId: vnet.id
        registrationEnabled: false
        location: 'global'
        tags: tags
      }
    ]
  }
}

// AVM module for Blob Private Endpoint with private DNS zone
module vaultPrivateEndpoint 'br/public:avm/res/network/private-endpoint:0.11.0' = {
  name: 'private-endpoint-vault-deployment'
  params: {
    name: 'private-endpoint-vault'
    location: location
    tags: tags
    subnetResourceId: '${vnet.id}/subnets/${subnetName}'
    privateLinkServiceConnections: [
      {
        name: 'pl-vault'
        properties: {
          privateLinkServiceId: keyVaultResource.id
          groupIds: [
            'vault'
          ]
        }
      }
    ]
    customDnsConfigs: []
    // Creates private DNS zone and links
    privateDnsZoneGroup: {
      name: 'vaultPrivateDnsZoneGroup'
      privateDnsZoneGroupConfigs: [
        {
          name: 'vaultARecord'
          privateDnsZoneResourceId: privateDnsZoneVaultDeployment.outputs.resourceId
        }
      ]
    }
  }
}

