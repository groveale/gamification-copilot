// NSG for APIM subnet
var apimNsgName = 'nsg-apim-${applicationName}'
resource apimSubnetNsg 'Microsoft.Network/networkSecurityGroups@2023-05-01' = {
  name: apimNsgName
  location: location
  tags: tags
  properties: {
    securityRules: [
      // Allow all outbound (customize as needed)
      {
        name: 'AllowAllOutbound'
        properties: {
          priority: 100
          protocol: '*'
          access: 'Allow'
          direction: 'Outbound'
          sourceAddressPrefix: '*'
          sourcePortRange: '*'
          destinationAddressPrefix: '*'
          destinationPortRange: '*'
        }
      }
      // Allow all inbound (customize as needed)
      {
        name: 'AllowAllInbound'
        properties: {
          priority: 200
          protocol: '*'
          access: 'Allow'
          direction: 'Inbound'
          sourceAddressPrefix: '*'
          sourcePortRange: '*'
          destinationAddressPrefix: '*'
          destinationPortRange: '*'
        }
      }
    ]
  }
}
@description('The location for all resources')
param location string = resourceGroup().location

@description('The runtime stack for the function app')
@allowed(['dotnet-isolated', 'node', 'python', 'java'])
param runtime string = 'dotnet-isolated'

@description('Storage account type')
@allowed(['Standard_LRS', 'Standard_GRS', 'Standard_RAGRS', 'Premium_LRS'])
param storageAccountType string = 'Standard_LRS'


@description('Virtual Network address prefix')
param vnetAddressPrefix string = '10.0.0.0/16'

@description('Function App subnet address prefix')
param functionSubnetAddressPrefix string = '10.0.1.0/24'

@description('API Management subnet address prefix')
param apimSubnetAddressPrefix string = '10.0.2.0/24'

var vnetName = 'vnet-${applicationName}'
var functionSubnetName = 'snet-function'
var apimSubnetName = 'snet-apim'
// Virtual Network for Flex Consumption VNet integration (with APIM subnet)
resource virtualNetwork 'Microsoft.Network/virtualNetworks@2023-05-01' = {
  name: vnetName
  location: location
  tags: tags
  properties: {
    addressSpace: {
      addressPrefixes: [
        vnetAddressPrefix
      ]
    }
    subnets: [
      {
        name: functionSubnetName
        properties: {
          addressPrefix: functionSubnetAddressPrefix
          delegations: [
            {
              name: 'delegation'
              properties: {
                serviceName: 'Microsoft.Web/serverFarms'
              }
            }
          ]
          serviceEndpoints: [
            {
              service: 'Microsoft.KeyVault'
            }
          ]
        }
      }
      {
        name: apimSubnetName
        properties: {
          addressPrefix: apimSubnetAddressPrefix
          networkSecurityGroup: {
            id: apimSubnetNsg.id
          }
        }
      }
    ]
  }
}

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

@description('Inactive days for the function app settings')
param inactivityDays string

@description('Object ID of the service account to grant Key Vault access - runs the flows')
param userObjectId string

@description('API Management SKU')
@allowed(['Developer', 'Basic', 'Standard', 'Premium'])
param apimSku string = 'Developer'

@description('API Management publisher email')
param apimPublisherEmail string

@description('API Management publisher name')
param apimPublisherName string

@description('Private endpoint subnet address prefix')
param privateEndpointSubnetAddressPrefix string = '10.0.3.0/24'

var storageAccountName = 'store${applicationName}'
var appServicePlanName = 'asp-${applicationName}'
var keyVaultName = 'kv-${applicationName}'
var functionAppName = 'func-${applicationName}'
var appInsightsName = 'appi-${applicationName}'
var apimName = 'apim-${applicationName}'


// (Removed: Network Security Groups and Virtual Network for Flex Consumption)

// Application Insights
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  kind: 'web'
  tags: tags
  properties: {
    Application_Type: 'web'
    RetentionInDays: 90
  }
}

// Storage Account with VNet restrictions
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
    allowBlobPublicAccess: false
    publicNetworkAccess: 'Enabled' // Changed from 'Disabled' to support VNet access
    networkAcls: {
      bypass: 'AzureServices'
      defaultAction: 'Deny'
      virtualNetworkRules: [
        {
          id: resourceId('Microsoft.Network/virtualNetworks/subnets', vnetName, functionSubnetName)
          action: 'Allow'
        }
      ]
    }
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
  dependsOn: [
    virtualNetwork
  ]
}

// Key Vault with VNet restrictions
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
    networkAcls: {
      bypass: 'AzureServices'
      defaultAction: 'Deny'
      virtualNetworkRules: [
        {
          id: resourceId('Microsoft.Network/virtualNetworks/subnets', vnetName, functionSubnetName)
          ignoreMissingVnetServiceEndpoint: false
        }
      ]
    }
  }
  dependsOn: [
    virtualNetwork
  ]
}

resource encryptionKeySecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: keyVaultEncryptionKeySecretName
  properties: {
    value: encryptionKey
  }
}

// App Service Plan (Flex Consumption)
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  tags: tags
  sku: {
    name: 'FC1'
    tier: 'FlexConsumption'
  }
  kind: 'functionapp'
  properties: {
    maximumElasticWorkerCount: 1
    perSiteScaling: false
    isSpot: false
  }
}

// Storage connection strings
var storageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'

// Function App (Flex Consumption with VNet integration)
resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: functionAppName
  location: location
  tags: tags
  kind: 'functionapp'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlan.id
    virtualNetworkSubnetId: resourceId('Microsoft.Network/virtualNetworks/subnets', vnetName, functionSubnetName)
    siteConfig: {
      appSettings: [
        {
          name: 'APPINSIGHTS_INSTRUMENTATIONKEY'
          value: appInsights.properties.InstrumentationKey
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
        {
          name: 'ApplicationInsightsAgent_EXTENSION_VERSION'
          value: '~3'
        }
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
          value: 'https://${storageAccount.name}.table.${environment().suffixes.storage}/'
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
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      http20Enabled: true
      use32BitWorkerProcess: false
    }
    httpsOnly: true
    clientAffinityEnabled: false
    publicNetworkAccess: 'Enabled'
  }
}

// API Management
resource apiManagement 'Microsoft.ApiManagement/service@2023-05-01-preview' = {
  name: apimName
  location: location
  tags: tags
  sku: {
    name: apimSku
    capacity: 1
  }
  properties: {
    publisherEmail: apimPublisherEmail
    publisherName: apimPublisherName
    virtualNetworkType: 'External'
    virtualNetworkConfiguration: {
      subnetResourceId: resourceId('Microsoft.Network/virtualNetworks/subnets', vnetName, apimSubnetName)
    }
    customProperties: {
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls10': 'False'
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Tls11': 'False'
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Protocols.Ssl30': 'False'
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls10': 'False'
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Tls11': 'False'
      'Microsoft.WindowsAzure.ApiManagement.Gateway.Security.Backend.Protocols.Ssl30': 'False'
    }
  }
  dependsOn: [
    virtualNetwork
  ]
}

// API Management Backend pointing to Function App
resource apimBackend 'Microsoft.ApiManagement/service/backends@2023-05-01-preview' = {
  name: '${functionAppName}-backend'
  parent: apiManagement
  properties: {
    description: 'Function App Backend'
    url: 'https://${functionApp.properties.defaultHostName}'
    protocol: 'http'
    resourceId: '${environment().resourceManager}${functionApp.id}'
    credentials: {
      header: {
        'x-functions-key': ['{{function-key}}']
      }
    }
  }
}

// API Management API
resource apimApi 'Microsoft.ApiManagement/service/apis@2023-05-01-preview' = {
  name: 'copilot-gamification-api'
  parent: apiManagement
  properties: {
    displayName: 'Copilot Gamification API'
    description: 'API for M365 Copilot Gamification Functions'
    path: 'api'
    protocols: ['https']
    serviceUrl: 'https://${functionApp.properties.defaultHostName}'
    subscriptionRequired: true
    subscriptionKeyParameterNames: {
      header: 'Ocp-Apim-Subscription-Key'
      query: 'subscription-key'
    }
  }
}

// API Management Operations for webhook endpoints
resource apimOperationWebhook 'Microsoft.ApiManagement/service/apis/operations@2023-05-01-preview' = {
  name: 'receive-events'
  parent: apimApi
  properties: {
    displayName: 'Receive Events'
    method: 'POST'
    urlTemplate: '/ReceiveEvents'
    description: 'Endpoint to receive M365 webhook events'
    request: {
      description: 'M365 webhook payload'
      headers: [
        {
          name: 'Content-Type'
          type: 'string'
          required: true
          values: ['application/json']
        }
      ]
    }
    responses: [
      {
        statusCode: 200
        description: 'Success'
      }
    ]
  }
}

// API Management Policy for the webhook operation
resource apimOperationPolicy 'Microsoft.ApiManagement/service/apis/operations/policies@2023-05-01-preview' = {
  name: 'policy'
  parent: apimOperationWebhook
  properties: {
    value: '''
    <policies>
      <inbound>
        <base />
        <set-backend-service backend-id="${functionAppName}-backend" />
        <rewrite-uri template="/api/ReceiveEvents" />
      </inbound>
      <backend>
        <base />
      </backend>
      <outbound>
        <base />
      </outbound>
      <on-error>
        <base />
      </on-error>
    </policies>
    '''
    format: 'xml'
  }
}

// Role assignments
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

// Outputs
output functionAppName string = functionApp.name
output functionAppHostName string = functionApp.properties.defaultHostName
output functionAppPrincipalId string = enableManagedIdentity ? functionApp.identity.principalId : ''
output storageAccountName string = storageAccount.name
output keyVaultName string = keyVault.name
output keyVaultUri string = keyVault.properties.vaultUri
output resourceGroupName string = resourceGroup().name
output appInsightsName string = appInsights.name
output appInsightsInstrumentationKey string = appInsights.properties.InstrumentationKey
output appInsightsConnectionString string = appInsights.properties.ConnectionString
output apiManagementName string = apiManagement.name
output apiManagementGatewayUrl string = apiManagement.properties.gatewayUrl
output vnetName string = virtualNetwork.name
output apiManagementPublicIP string = apiManagement.properties.publicIPAddresses[0]
