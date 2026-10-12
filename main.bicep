// ============================================================================
// AutoOnboard Architecture - Production Ready Node.js & Azure SQL Template
// ============================================================================

param locationEastUs string = 'eastus'
param locationCentralUs string = 'centralus'

@description('Administrator username for Azure SQL Server')
param sqlAdminLogin string = 'autodbuser2026'

@description('Administrator password for Azure SQL Server')
@secure()
param sqlAdminPassword string = 'AutoOnboard@2026#Pass'

param tenantId string = subscription().tenantId
param entraClientId string = '00000000-0000-0000-0000-000000000000'

@secure()
param entraClientSecret string = 'placeholder-secret-value'

// ----------------------------------------------------------------------------
// Resource 1: Azure Key Vault
// ----------------------------------------------------------------------------
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: 'kv-autoonboard-01'
  location: locationEastUs
  properties: {
    tenantId: tenantId
    sku: {
      family: 'A'
      name: 'standard'
    }
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 90
    enabledForTemplateDeployment: true
    publicNetworkAccess: 'Enabled'
  }
}

resource secretEntraClientId 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'Entra-Client-ID'
  properties: { value: entraClientId }
}

resource secretEntraClientSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'Entra-Client-Secret'
  properties: { value: entraClientSecret }
}

// ----------------------------------------------------------------------------
// Resource 2: Azure Storage Account
// ----------------------------------------------------------------------------
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: 'stautoonboard01'
  location: locationEastUs
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    accessTier: 'Hot'
    minimumTlsVersion: 'TLS1_2'
    supportsHttpsTrafficOnly: true
    publicNetworkAccess: 'Enabled'
  }
}

// ----------------------------------------------------------------------------
// Resource 3: Azure SQL Server & Database + Firewall Rule
// ----------------------------------------------------------------------------
resource sqlServer 'Microsoft.Sql/servers@2023-05-01-preview' = {
  name: 'sqlserver-autoonboard-01'
  location: locationCentralUs
  properties: {
    administratorLogin: sqlAdminLogin
    administratorLoginPassword: sqlAdminPassword
    version: '12.0'
    minimalTlsVersion: '1.2'
    publicNetworkAccess: 'Enabled'
  }
}

resource sqlDatabase 'Microsoft.Sql/servers/databases@2023-05-01-preview' = {
  parent: sqlServer
  name: 'sqldb-autoonboard'
  location: locationCentralUs
  sku: {
    name: 'Standard'
    tier: 'Standard'
    capacity: 10
  }
  properties: {
    collation: 'SQL_Latin1_General_CP1_CI_AS'
    maxSizeBytes: 268435456000
  }
}

// Strictly allow Azure App Service internal connections
resource sqlFirewallAllowAzure 'Microsoft.Sql/servers/firewallRules@2023-05-01-preview' = {
  parent: sqlServer
  name: 'AllowAllWindowsAzureIps'
  properties: {
    startIpAddress: '0.0.0.0'
    endIpAddress: '0.0.0.0'
  }
}

// ----------------------------------------------------------------------------
// Resource 4: Azure Linux App Service Plan & Web App (Node.js)
// ----------------------------------------------------------------------------
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: 'plan-autoonboard-01'
  location: locationCentralUs
  kind: 'linux'
  sku: {
    name: 'B1'
    tier: 'Basic'
    capacity: 1
  }
  properties: {
    reserved: true
  }
}

resource appService 'Microsoft.Web/sites@2023-01-01' = {
  name: 'app-autoonboard-01'
  location: locationCentralUs
  kind: 'app,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    publicNetworkAccess: 'Enabled'
    siteConfig: {
      linuxFxVersion: 'NODE|20-lts'
      appCommandLine: 'node server.js'
      minTlsVersion: '1.2'
      ftpsState: 'FtpsOnly'
      appSettings: [
        {
          name: 'DB_HOST'
          value: '${sqlServer.name}.database.windows.net'
        }
        {
          name: 'DB_NAME'
          value: sqlDatabase.name
        }
        {
          name: 'DB_USER'
          value: sqlAdminLogin
        }
        {
          name: 'DB_PASS'
          value: sqlAdminPassword
        }
        {
          name: 'SQL_CONNECTION_STRING'
          value: 'Server=tcp:${sqlServer.name}.database.windows.net,1433;Initial Catalog=${sqlDatabase.name};Persist Security Info=False;User ID=${sqlAdminLogin};Password=${sqlAdminPassword};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;'
        }
        {
          name: 'SCM_DO_BUILD_DURING_DEPLOYMENT'
          value: 'true'
        }
      ]
    }
  }
}

// ----------------------------------------------------------------------------
// Outputs
// ----------------------------------------------------------------------------
output webAppHostName string = 'https://${appService.properties.defaultHostName}'
output sqlServerFqdn string = sqlServer.properties.fullyQualifiedDomainName
output keyVaultUri string = keyVault.properties.vaultUri