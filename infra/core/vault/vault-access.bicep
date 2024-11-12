param principalID string
param roleDefinitionID string
param keyVaultName string

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' existing = {
  name: keyVaultName
}

// Allow access from API to storage account using a managed identity
resource vaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, principalID, roleDefinitionID)
  scope: keyVault
  properties: {
    roleDefinitionId: roleDefinitionID
    principalId: principalID
    principalType: 'ServicePrincipal' // Workaround for https://learn.microsoft.com/en-us/azure/role-based-access-control/role-assignments-template#new-service-principal
  }
}

output ROLE_ASSIGNMENT_NAME string = vaultRoleAssignment.name
