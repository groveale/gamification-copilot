# Secure Azure Function App Deployment with API Management

This deployment creates a secure, production-ready infrastructure for your Azure Function App with the following security features:

## Security Architecture

### Network Security
- **VNet Integration**: Function App deployed within a private virtual network
- **Private Access Only**: Function App has public network access disabled
- **API Management Gateway**: Only APIM has public access, acting as a secure proxy
- **Network Security Groups**: Restricts inbound traffic to M365 service tags only

### Data Security
- **Key Vault Integration**: Secrets stored securely in Azure Key Vault with VNet restrictions
- **Storage Account Security**: Private access with VNet service endpoints
- **Managed Identity**: Function App uses system-assigned managed identity for secure authentication
- **Encryption**: All data encrypted at rest and in transit

### Network Topology

```
M365 Webhook (Service Tag: M365ManagementActivityApiWebhook) ────┐
                                                                 ↓ HTTPS (Port 443)
Power Automate (Service Tag: PowerPlatformInfra) ───────────────┤
                                                                 ↓ [Network Security Group]
                                                                 ↓ (DENY all Internet traffic)
                                                                 ↓ (ALLOW only specific service tags)
                                                              [API Management Subnet] (10.0.2.0/24)
                                                                 ↓ Private HTTPS
                                                              [Function App Subnet] (10.0.1.0/24)
                                                                 ↓ ↑ Outbound HTTPS calls
                                                                 ↓ ↑ M365ManagementActivityApi
                                                                 ↓ ↑ AzureActiveDirectory
                                                                 ↓ ↑ AzureCloud (Microsoft Graph)
                                                              [Azure Services] (Storage, Key Vault)
```

## Infrastructure Components

### Core Components
- **Azure Function App**: Hosted on Premium App Service Plan with VNet integration
- **API Management**: Developer/Basic tier with external VNet integration
- **Virtual Network**: Segmented subnets for different components
- **Network Security Groups**: Firewall rules for traffic filtering

### Supporting Services
- **Storage Account**: With VNet restrictions and managed identity access
- **Key Vault**: Secure secret storage with RBAC and VNet access
- **Application Insights**: Monitoring and telemetry
- **App Service Plan**: Premium tier (required for VNet integration)

## Subnet Architecture

| Subnet | Address Range | Purpose | Restrictions |
|--------|--------------|---------|--------------|
| Function Subnet | 10.0.1.0/24 | Function App VNet integration | Delegated to Microsoft.Web/serverFarms, NSG with outbound rules |
| APIM Subnet | 10.0.2.0/24 | API Management | NSG with M365 and Power Platform inbound rules |
| Private Endpoint Subnet | 10.0.3.0/24 | Future private endpoints | Private endpoint policies disabled |

## Security Rules

### Network Security Group Rules

#### APIM Subnet (Inbound Rules)
- **Allow M365 Management Activity API**: Inbound HTTPS from M365ManagementActivityApiWebhook service tag only
- **Allow Power Automate**: Inbound HTTPS from PowerPlatformInfra service tag only (covers Power Automate/Power Platform)
- **Allow APIM Management**: Management plane access for API Management service
- **Allow Function Communication**: Outbound from APIM to Function subnet
- **Deny All Other Traffic**: Default deny rule blocks all other inbound traffic (including general Internet)

#### Function Subnet (Outbound Rules)
- **Allow M365 Management Activity API**: Outbound HTTPS to M365ManagementActivityApi service tag
- **Allow Azure Active Directory**: Outbound HTTPS to AzureActiveDirectory service tag (required for M365 API authentication)
- **Allow Azure Services**: Outbound HTTPS to AzureCloud service tag (Storage, Key Vault, Microsoft Graph for SharePoint)
- **Deny All Other Traffic**: Default deny rule blocks all other outbound traffic

### IP Security Restrictions
- Function App only accepts traffic from APIM subnet
- All other traffic is denied at the application level

## Deployment Instructions

### Prerequisites
1. Azure PowerShell module installed
2. Azure account with Contributor rights to subscription
3. Required parameter values (see below)

### Required Parameters
Before deployment, gather these values:

- **ResourceGroupName**: Target resource group name
- **ApplicationName**: Unique name for your application (used in resource naming)
- **EncryptionKey**: Strong encryption key for data encryption
- **TenantId**: Your Azure AD tenant ID
- **AuthGuid**: Authentication GUID for your application
- **SpoSiteId**: SharePoint Online site ID
- **SpoListId**: SharePoint Online list ID
- **UserObjectId**: Your user object ID for Key Vault access
- **ApimPublisherEmail**: Email for API Management publisher
- **ApimPublisherName**: Name for API Management publisher

### Deployment Steps

1. **Login to Azure**:
   ```powershell
   Connect-AzAccount
   ```

2. **Choose your deployment method**:

   **Option A: Using Parameters File (Recommended)**
   
   First, copy and customize the parameters file:
   ```powershell
   # Copy the template parameters file
   Copy-Item "main-secure.parameters.json" "my-deployment.parameters.json"
   
   # Edit the parameters file with your values
   notepad my-deployment.parameters.json
   ```
   
   Then deploy:
   ```powershell
   .\deploy-secure.ps1 -ResourceGroupName "rg-copilot-gamify-prod" `
                       -ParametersFile "my-deployment.parameters.json"
   ```

   **Option B: Using Inline Parameters**
   ```powershell
   .\deploy-secure.ps1 -ResourceGroupName "rg-copilot-gamify-prod" `
                       -ApplicationName "copilotgamify" `
                       -EncryptionKey "YourStrongEncryptionKey123!" `
                       -TenantId "your-tenant-id" `
                       -AuthGuid "your-auth-guid" `
                       -SpoSiteId "your-spo-site-id" `
                       -SpoListId "your-spo-list-id" `
                       -UserObjectId "your-user-object-id" `
                       -ApimPublisherEmail "admin@yourcompany.com" `
                       -ApimPublisherName "Your Company"
   ```

3. **Deploy Function App Code**:
   After infrastructure deployment, deploy your function app code:
   ```powershell
   # Navigate to your function app source
   cd ..\src
   
   # Publish to Azure
   func azure functionapp publish func-copilotgamify --dotnet-isolated
   ```

## Post-Deployment Configuration

### 1. Configure M365 Webhook
Update your M365 Management Activity API webhook URL to:
```
https://apim-copilotgamify.azure-api.net/api/ReceiveEvents
```

### 2. Get API Management Subscription Key
1. Navigate to API Management in Azure Portal
2. Go to Subscriptions
3. Copy the primary or secondary key
4. Use this key in the `Ocp-Apim-Subscription-Key` header for webhook calls

### 3. Configure Power Automate Integration
Power Automate can call your APIM endpoints using the **HTTP** connector:

**URL Format**: `https://apim-copilotgamify.azure-api.net/api/{function-name}`

**Required Headers**:
```
Ocp-Apim-Subscription-Key: {your-subscription-key}
Content-Type: application/json
```

**Example Power Automate HTTP Action**:
- **Method**: POST/GET (depending on your function)
- **URI**: `https://apim-copilotgamify.azure-api.net/api/GetTodaysCopilotUsage`
- **Headers**: 
  - `Ocp-Apim-Subscription-Key`: `your-subscription-key-here`
  - `Content-Type`: `application/json`
- **Body**: JSON payload as required by your function

### 4. Common Power Automate Use Cases
With this secure architecture, you can build Power Automate flows for:

- **📊 Daily Reports**: Scheduled flows calling `GetTodaysCopilotUsage`
- **🏆 Leaderboards**: Flows calling `GetUsersWithStreak` for gamification
- **📧 Notifications**: Flows calling `GetNotifications` and sending emails
- **⚠️ Inactive User Alerts**: Flows calling `GetInactiveUsers` for HR teams
- **📈 Analytics Dashboard**: Multi-step flows aggregating usage data

**Security Benefits**:
- ✅ Power Automate cannot directly access your Function App
- ✅ All calls are logged and monitored through APIM
- ✅ Rate limiting and throttling policies apply
- ✅ Subscription keys can be rotated without changing flows

### 5. Verify Security
- Confirm Function App shows "Access restricted" in Azure Portal
- Verify APIM gateway URL is accessible from internet and Power Automate
- Test direct Function App URL is not accessible (should be blocked)
- Test Power Automate flows can successfully call APIM endpoints

## Monitoring and Troubleshooting

### Application Insights
Monitor your application through Application Insights:
- Function execution logs
- Performance metrics
- Dependency tracking
- Custom telemetry

### Network Diagnostics
Use Azure Network Watcher for:
- Connection troubleshooting
- Flow logs analysis
- Security group diagnostics

### Common Issues

1. **Function App not accessible through APIM**:
   - Check NSG rules allow traffic between subnets
   - Verify APIM backend configuration
   - Confirm Function App IP restrictions

2. **Webhook calls failing**:
   - Verify NSG allows M365ManagementActivityApi service tag
   - Check APIM subscription key is included in requests
   - Confirm SSL/TLS certificates are valid

3. **Storage/Key Vault access issues**:
   - Verify managed identity has correct RBAC roles
   - Check VNet service endpoints are configured
   - Confirm network ACLs allow Function subnet

## Security Considerations

### Best Practices Implemented
- ✅ **Zero Trust Network**: All traffic explicitly allowed via service tags only
- ✅ **No Internet Access**: General internet traffic completely blocked at NSG level
- ✅ **Service Tag Filtering**: Only M365ManagementActivityApiWebhook and PowerPlatformInfra service tags allowed (inbound)
- ✅ **Principle of Least Privilege**: Minimal required permissions and network access
- ✅ **Defense in Depth**: Multiple security layers (NSG, APIM, Function App restrictions)
- ✅ **Encrypted Communications**: TLS 1.2+ everywhere
- ✅ **Secret Management**: Centralized in Key Vault
- ✅ **Identity-Based Access**: Managed identities used

### Additional Recommendations
- Consider Azure Private DNS for name resolution
- Implement Azure Firewall for additional filtering
- Use Azure Security Center recommendations
- Regular security assessments and penetration testing
- Monitor with Azure Sentinel for advanced threat detection

## Cost Optimization

### Resource Sizing
- **App Service Plan**: Start with P1v2, scale as needed
- **API Management**: Developer tier for non-production, Standard+ for production
- **Storage Account**: Standard_LRS sufficient for most scenarios

### Monitoring Costs
- Set up budget alerts
- Use Azure Advisor recommendations
- Monitor Application Insights data ingestion

## Disaster Recovery

### Backup Strategy
- Function App code in source control
- Infrastructure as Code (this Bicep template)
- Key Vault secrets backup
- Regular testing of deployment process

### High Availability
- Consider multi-region deployment for critical workloads
- Use Azure Traffic Manager for global load balancing
- Implement Circuit Breaker pattern in function code
