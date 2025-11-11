@description('The location for all resources')
param location string = resourceGroup().location

@description('The runtime stack for the function app')
@allowed(['dotnet-isolated', 'node', 'python', 'java'])
param runtime string = 'dotnet-isolated'

@description('Storage account type')
@allowed(['Standard_LRS', 'Standard_GRS', 'Standard_RAGRS', 'Premium_LRS'])
param storageAccountType string = 'Standard_LRS'

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
param storagePublicNetworkAccess bool = false

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

@description('Maximum instance count for Flex Consumption')
param maximumInstanceCount int = 100

@description('Instance memory in MB for Flex Consumption (2048, 4096)')
@allowed([2048, 4096])
param instanceMemoryMB int = 2048

@description('Allowed IP addresses for inbound access restrictions')
param allowedIpAddresses array = []

@description('Allowed service tags for inbound access restrictions')
param allowedServiceTags array = [
  'AzureCloud'
]

@description('Allowed virtual networks for inbound access restrictions')
param allowedVirtualNetworks array = []

var storageAccountName = 'store${applicationName}'
var functionAppName = 'func-${applicationName}'
var appInsightsName = 'appi-${applicationName}'
var keyVaultName = 'kv-${applicationName}'

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
    allowSharedKeyAccess: false // Disable shared key access to force managed identity
    defaultToOAuthAuthentication: true
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
    networkAcls: {
      defaultAction: storagePublicNetworkAccess ? 'Allow' : 'Deny'
      bypass: 'AzureServices'
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

// Flex Consumption plan
resource flexConsumptionPlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: 'flex-${applicationName}'
  location: location
  tags: tags
  sku: {
    name: 'FC1'
    tier: 'FlexConsumption'
  }
  properties: {
    reserved: runtime == 'node' || runtime == 'python'
  }
}

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  dependsOn: enableAppInsights ? [appInsights] : []
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: flexConsumptionPlan.id
    functionAppConfig: {
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storageAccount.properties.primaryEndpoints.blob}deployments'
          authentication: {
            type: 'SystemAssignedIdentity'
          }
        }
      }
      scaleAndConcurrency: {
        maximumInstanceCount: maximumInstanceCount
        instanceMemoryMB: instanceMemoryMB
      }
      runtime: {
        name: runtime
        version: runtime == 'dotnet-isolated' ? '8.0' : (runtime == 'node' ? '20' : (runtime == 'python' ? '3.11' : '17'))
      }
    }
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
        // Managed Identity configuration for storage
        {
          name: 'AzureWebJobsStorage__blobServiceUri'
          value: storageAccount.properties.primaryEndpoints.blob
        }
        {
          name: 'AzureWebJobsStorage__queueServiceUri'
          value: storageAccount.properties.primaryEndpoints.queue
        }
        {
          name: 'AzureWebJobsStorage__tableServiceUri'
          value: storageAccount.properties.primaryEndpoints.table
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING__credential'
          value: 'managedidentity'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING__blobServiceUri'
          value: storageAccount.properties.primaryEndpoints.blob
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
          value: storageAccount.properties.primaryEndpoints.table
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

// Role definitions
var keyVaultSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'
var storageBlobDataContributorRoleId = 'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var storageQueueDataContributorRoleId = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'

// User access to Key Vault
resource userKeyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(keyVault.id, userObjectId, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalId: userObjectId
    principalType: 'User'
  }
}

// Function App managed identity access to Key Vault
resource keyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (enableManagedIdentity) {
  name: guid(keyVault.id, functionApp.id, keyVaultSecretsUserRoleId)
  scope: keyVault
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalId: functionApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// Function App managed identity access to Storage - Blob
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

// Function App managed identity access to Storage - Table
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

// Function App managed identity access to Storage - Queue
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
