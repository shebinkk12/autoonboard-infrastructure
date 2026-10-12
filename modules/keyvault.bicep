@description('Azure region for Key Vault')
param location string

@description('Key Vault name')
param keyVaultName string

@description('Azure Tenant ID')
param tenantId string

@description('Entra ID App Client ID')
@secure()
param clientId string

@description('Entra ID App Client Secret')
@secure()
param clientSecret string

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  properties: {
    tenantId: tenantId
    sku: {
      family: 'A'
      name: 'standard'
    }
    enableRbacAuthorization: true
    enabledForDeployment: false
    enabledForTemplateDeployment: true
    enabledForDiskEncryption: false
  }
}

resource secretTenantId 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'Entra-Tenant-ID'
  properties: {
    value: tenantId
  }
}

resource secretClientId 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'Entra-Client-ID'
  properties: {
    value: clientId
  }
}

resource secretClientSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'Entra-Client-Secret'
  properties: {
    value: clientSecret
  }
}

output keyVaultId string = keyVault.id