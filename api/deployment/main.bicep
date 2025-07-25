@description('The location for all resources')
param location string = resourceGroup().location

@description('The runtime stack for the function app')
@allowed(['dotnet-isolated', 'node', 'python', 'java'])
param runtime string = 'dotnet-isolated'

@description('Storage account type')
@allowed(['Standard_LRS', 'Standard_GRS', 'Standard_RAGRS', 'Premium_LRS'])
param storageAccountType string = 'Standard_LRS'

@description('App Service Plan SKU')
@allowed(['Y1', 'B1', 'EP1', 'EP2', 'EP3', 'P1v2', 'P2v2', 'P3v2'])
param appServicePlanSku string = 'Y1'

@description('Application name for resource naming')
param applicationName string

@description('Enable managed identity for the function app')
param enableManagedIdentity bool = true

@description('Key Vault SKU')
@allowed(['standard', 'premium'])
param keyVaultSku string = 'standard'

@secure()
@description('The encryption key to be stored in Key Vault')
param encryptionKey string

@description('Key Vault encryption key secret name')
param keyVaultEncryptionKeySecretName string = 'email-encryption-key'

@description('Entra Tenant Id')
param tenantId string

@description('Authentication GUID')
param authGuid string

@description('SharePoint Online Site ID')
param spoSiteId string

@description('SharePoint Online Field Name')
param spoFieldName string = 'UPN'

@description('SharePoint Online List ID')
param spoListId string

@description('Tags to apply to all resources')
param tags object = {
  Application: applicationName
  ManagedBy: 'Bicep'
}

@description('Enable public network access for storage account')
param storagePublicNetworkAccess bool = true

@description('Enable blob public access for storage account')
param allowBlobPublicAccess bool = false

@description('Enable Application Insights')
param enableAppInsights bool = true

@description('App Insights location (defaults to function app location)')
param appInsightsLocation string = location

@description('Inactive days for the function app settings')
param inactivityDays string

@description('Object ID of the service account to grant Key Vault access - runs the flows')
param userObjectId string

var storageAccountName = 'store${applicationName}'
var appServicePlanName = 'asp-${applicationName}'
var keyVaultName = 'kv-${applicationName}'
var functionAppName = 'func-${applicationName}'
var appInsightsName = 'appi-${applicationName}'

resource appInsights 'Microsoft.Insights/components@2020-02-02' = if (enableAppInsights) {
  name: appInsightsName
  location: appInsightsLocation
  kind: 'web'
  tags: tags
  properties: {
    Application_Type: 'web'
    RetentionInDays: 90
  }
}

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageAccountName
  location: location
  tags: tags
  sku: {
    name: storageAccountType
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: allowBlobPublicAccess
    publicNetworkAccess: storagePublicNetworkAccess ? 'Enabled' : 'Disabled'
    encryption: {
      services: {
        file: {
          keyType: 'Account'
          enabled: true
        }
        blob: {
          keyType: 'Account'
          enabled: true
        }
      }
      keySource: 'Microsoft.Storage'
    }
  }
}

resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  tags: tags
  properties: {
    sku: {
      family: 'A'
      name: keyVaultSku
    }
    tenantId: tenant().tenantId
    enabledForDeployment: false
    enabledForDiskEncryption: false
    enabledForTemplateDeployment: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
    enableRbacAuthorization: true
    publicNetworkAccess: 'Enabled'
  }
}

resource encryptionKeySecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: keyVaultEncryptionKeySecretName
  properties: {
    value: encryptionKey
  }
}

resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  tags: tags
  sku: {
    name: appServicePlanSku
  }
  properties: {
    reserved: runtime == 'node' || runtime == 'python'
  }
}

// Add these variables to generate the connection strings
var storageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'


resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  dependsOn: enableAppInsights ? [appInsights] : []
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      appSettings: [
        {
          name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
          value: enableAppInsights ? appInsights.properties.InstrumentationKey : ''
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: enableAppInsights ? appInsights.properties.ConnectionString : ''
        }
        {
          name: 'ApplicationInsightsAgent_EXTENSION_VERSION'
          value: '~3'
        }
        // Add these required connection strings for Consumption plan
        {
          name: 'AzureWebJobsStorage'
          value: storageConnectionString
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: storageConnectionString
        }
        {
          name: 'AzureWebJobsStorage__credential'
          value: 'managedidentity'
        }
        {
          name: 'AzureWebJobsStorage__accountName'
          value: storageAccount.name
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING__credential'
          value: 'managedidentity'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING__accountName'
          value: storageAccount.name
        }
        {
          name: 'WEBSITE_CONTENTSHARE'
          value: toLower(functionAppName)
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: runtime
        }
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
        {
          name: 'tenantId'
          value: tenantId
        }
        {
          name: 'AuthGuid'
          value: authGuid
        }
        {
          name: 'StorageAccountName'
          value: storageAccount.name
        }
        {
          name: 'StorageAccountUri'
          value: 'https://${storageAccount.name}.table.core.windows.net/'
        }
        {
          name: 'KeyVault:Url'
          value: keyVault.properties.vaultUri
        }
        {
          name: 'KeyVault:EncryptionKeySecretName'
          value: keyVaultEncryptionKeySecretName
        }
        {
          name: 'SPO:SiteId'
          value: spoSiteId
        }
        {
          name: 'SPO:FieldName'
          value: spoFieldName
        }
        {
          name: 'SPO:ListId'
          value: spoListId
        }
        {
          name: 'ReminderDays'
          value: inactivityDays
        }
      ]
      ftpsState: 'FtpsOnly'
      minTlsVersion: '1.2'
      http20Enabled: true
      use32BitWorkerProcess: false
      cors: {
        allowedOrigins: [
          'https://portal.azure.com'
        ]
      }
    }
    httpsOnly: true
    clientAffinityEnabled: false
  }
}

var keyVaultSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'
var storageBlobDataContributorRoleId = 'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var storageQueueDataContributorRoleId = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'

resource userKeyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, userObjectId, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalId: userObjectId
    principalType: 'User'
  }
}

resource keyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (enableManagedIdentity) {
  name: guid(keyVault.id, functionApp.id, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (enableManagedIdentity) {
  name: guid(storageAccount.id, functionApp.id, storageBlobDataContributorRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageBlobDataContributorRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageTableRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (enableManagedIdentity) {
  name: guid(storageAccount.id, functionApp.id, storageTableDataContributorRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageTableDataContributorRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource storageQueueRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (enableManagedIdentity) {
  name: guid(storageAccount.id, functionApp.id, storageQueueDataContributorRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageQueueDataContributorRoleId
    )
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

output functionAppName string = functionApp.name
output functionAppHostName string = functionApp.properties.defaultHostName
output functionAppPrincipalId string = enableManagedIdentity ? functionApp.identity.principalId : ''
output storageAccountName string = storageAccount.name
output keyVaultName string = keyVault.name
output keyVaultUri string = keyVault.properties.vaultUri
output resourceGroupName string = resourceGroup().name
output appInsightsName string = enableAppInsights ? appInsights.name : ''
output appInsightsInstrumentationKey string = enableAppInsights ? appInsights.properties.InstrumentationKey : ''
output appInsightsConnectionString string = enableAppInsights ? appInsights.properties.ConnectionString : ''
