@description('Specifies the name of the virtual network.')
param vNetName string

@description('Specifies the location.')
param location string = resourceGroup().location

@description('Specifies the name of the subnet for the Service Bus private endpoint.')
param peSubnetName string = 'private-endpoints-subnet'

@description('Specifies the name of the subnet for Function App virtual network integration.')
param appSubnetName string = 'app'

param tags object = {}

// Migrated to use AVM module instead of direct resource declaration
module virtualNetwork 'br/public:avm/res/network/virtual-network:0.6.1' = {
  name: 'vnet-deployment'
  params: {
    // Required parameters
    name: vNetName
    addressPrefixes: [
      '10.0.0.0/16'
    ]
    // Non-required parameters
    location: location
    tags: tags
    subnets: [
      {
        name: peSubnetName
        addressPrefix: '10.0.1.0/24'
        privateEndpointNetworkPolicies: 'Disabled'
        privateLinkServiceNetworkPolicies: 'Enabled'
        defaultOutboundAccess: false
      }
      {
        name: appSubnetName
        addressPrefix: '10.0.2.0/24'
        privateEndpointNetworkPolicies: 'Disabled'
        privateLinkServiceNetworkPolicies: 'Enabled'
        delegation: 'Microsoft.App/environments'
        // If defaultOutboundAccess is set to false on this subnet and the function app's managed identity is SystemAssigned, deploying the app package fails with this error:
        // InaccessibleStorageException: Failed to access storage account for deployment: BlobUploadFailedException: Failed to upload blob to storage account: Response status code does not indicate success: 404
        // defaultOutboundAccess: false
      }
    ]
  }
}

output peSubnetName string = peSubnetName
output peSubnetID string = '${virtualNetwork.outputs.resourceId}/subnets/${peSubnetName}'
output appSubnetName string = appSubnetName
output appSubnetID string = '${virtualNetwork.outputs.resourceId}/subnets/${appSubnetName}'
