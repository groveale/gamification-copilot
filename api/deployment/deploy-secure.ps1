<#
.SYNOPSIS
    Deploys a secure Azure Function App with VNet integration and API Management

.DESCRIPTION
    This script deploys a secure infrastructure for the Copilot Gamification system using:
    - Azure Function App with VNet integration (Premium plan)
    - API Management with External VNet integration
    - Network Security Groups with service tag restrictions
    - Key Vault with VNet restrictions
    - Storage Account with service endpoints
    - Application Insights

.PARAMETER ResourceGroupName
    The name of the resource group to create or use

.PARAMETER Location
    The Azure region to deploy to (e.g., 'East US')

.PARAMETER ApplicationName
    The base name for the application resources (required when not using parameters file)

.PARAMETER EncryptionKey
    The encryption key for data protection (required when not using parameters file)

.PARAMETER TenantId
    The Azure AD tenant ID (required when not using parameters file)

.PARAMETER AuthGuid
    The authentication GUID for the application (required when not using parameters file)

.PARAMETER SpoSiteId
    The SharePoint site ID (required when not using parameters file)

.PARAMETER SpoListId
    The SharePoint list ID (required when not using parameters file)

.PARAMETER UserObjectId
    The object ID of the user to grant Key Vault access (required when not using parameters file)

.PARAMETER ApimPublisherEmail
    The email for the APIM publisher (required when not using parameters file)

.PARAMETER ApimPublisherName
    The name for the APIM publisher (required when not using parameters file)

.PARAMETER InactivityDays
    Number of days before marking a user as inactive (optional, defaults to 30)

.PARAMETER ParametersFile
    Optional path to a JSON parameters file (e.g., main-secure.parameters.json)

.EXAMPLE
    # Deploy using parameters file
    .\deploy-secure.ps1 -ResourceGroupName "rg-copilot-secure" -ParametersFile ".\main-secure.parameters.json"

.EXAMPLE
    # Deploy using inline parameters
    .\deploy-secure.ps1 -ResourceGroupName "rg-copilot-secure" -Location "East US" -ApplicationName "copilot-gamify" -EncryptionKey "your-encryption-key" -TenantId "your-tenant-id" -AuthGuid "your-auth-guid" -SpoSiteId "your-site-id" -SpoListId "your-list-id" -UserObjectId "your-user-object-id" -ApimPublisherEmail "admin@contoso.com" -ApimPublisherName "Contoso Admin" -InactivityDays 30
#>

# Deploy Secure Azure Function App with APIM
# This script deploys a secure infrastructure with VNet integration

param(
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$false)]
    [string]$Location = "East US",
    
    [Parameter(Mandatory=$false)]
    [string]$ApplicationName,
    
    [Parameter(Mandatory=$false)]
    [string]$EncryptionKey,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$AuthGuid,
    
    [Parameter(Mandatory=$false)]
    [string]$SpoSiteId,
    
    [Parameter(Mandatory=$false)]
    [string]$SpoListId,
    
    [Parameter(Mandatory=$false)]
    [string]$UserObjectId,
    
    [Parameter(Mandatory=$false)]
    [string]$ApimPublisherEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$ApimPublisherName,
    
    [Parameter(Mandatory=$false)]
    [string]$InactivityDays = "30",
    
    [Parameter(Mandatory=$false)]
    [string]$ParametersFile = $null
)

# Set error handling
$ErrorActionPreference = "Stop"

Write-Host "Starting secure deployment of Copilot Gamification infrastructure..." -ForegroundColor Green

# Validate parameters
if (-not $ParametersFile) {
    # If no parameters file, validate that all required inline parameters are provided
    $requiredParams = @('ApplicationName', 'EncryptionKey', 'TenantId', 'AuthGuid', 'SpoSiteId', 'SpoListId', 'UserObjectId', 'ApimPublisherEmail', 'ApimPublisherName')
    $missingParams = @()
    
    foreach ($param in $requiredParams) {
        $value = Get-Variable -Name $param -ValueOnly -ErrorAction SilentlyContinue
        if (-not $value) {
            $missingParams += $param
        }
    }
    
    if ($missingParams.Count -gt 0) {
        Write-Host "ERROR: When not using a parameters file, the following parameters are required:" -ForegroundColor Red
        Write-Host "Missing parameters: $($missingParams -join ', ')" -ForegroundColor Red
        Write-Host "Use -ParametersFile to provide parameters via JSON file, or provide all required inline parameters." -ForegroundColor Yellow
        exit 1
    }
}

try {
    # Check if logged in to Azure
    $context = Get-AzContext
    if (!$context) {
        Write-Host "Please login to Azure first using Connect-AzAccount" -ForegroundColor Red
        exit 1
    }

    Write-Host "Current Azure Context: $($context.Account.Id)" -ForegroundColor Yellow

    # Create resource group if it doesn't exist
    $rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    if (!$rg) {
        Write-Host "Creating resource group: $ResourceGroupName" -ForegroundColor Yellow
        New-AzResourceGroup -Name $ResourceGroupName -Location $Location
    } else {
        Write-Host "Resource group $ResourceGroupName already exists" -ForegroundColor Green
    }

    # Deploy the infrastructure
    Write-Host "Deploying secure infrastructure template..." -ForegroundColor Yellow
    
    $deploymentName = "copilot-gamify-secure-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    
    # Choose deployment method based on whether parameters file is provided
    if ($ParametersFile) {
        Write-Host "Using parameters file: $ParametersFile" -ForegroundColor Yellow
        
        # Verify parameters file exists
        if (-not (Test-Path $ParametersFile)) {
            Write-Host "Parameters file not found: $ParametersFile" -ForegroundColor Red
            exit 1
        }
        
        $deployment = New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName -Name $deploymentName -TemplateFile ".\main-secure.bicep" -TemplateParameterFile $ParametersFile -Verbose
    } else {
        Write-Host "Using inline parameters..." -ForegroundColor Yellow
        
        $templateParams = @{
            'location' = $Location
            'applicationName' = $ApplicationName
            'encryptionKey' = $EncryptionKey
            'tenantId' = $TenantId
            'authGuid' = $AuthGuid
            'spoSiteId' = $SpoSiteId
            'spoListId' = $SpoListId
            'userObjectId' = $UserObjectId
            'apimPublisherEmail' = $ApimPublisherEmail
            'apimPublisherName' = $ApimPublisherName
            'inactivityDays' = $InactivityDays
        }

        $deployment = New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName -Name $deploymentName -TemplateFile ".\main-secure.bicep" -TemplateParameterObject $templateParams -Verbose
    }

    if ($deployment.ProvisioningState -eq "Succeeded") {
        Write-Host "Infrastructure deployment completed successfully!" -ForegroundColor Green
        
        # Output important information
        Write-Host "`n=== Deployment Outputs ===" -ForegroundColor Cyan
        Write-Host "Function App Name: $($deployment.Outputs.functionAppName.Value)" -ForegroundColor White
        Write-Host "API Management Gateway URL: $($deployment.Outputs.apiManagementGatewayUrl.Value)" -ForegroundColor White
        Write-Host "API Management Public IP: $($deployment.Outputs.apiManagementPublicIP.Value)" -ForegroundColor White
        Write-Host "Key Vault Name: $($deployment.Outputs.keyVaultName.Value)" -ForegroundColor White
        Write-Host "Storage Account Name: $($deployment.Outputs.storageAccountName.Value)" -ForegroundColor White
        Write-Host "VNet Name: $($deployment.Outputs.vnetName.Value)" -ForegroundColor White

        Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
        Write-Host "1. Deploy your Function App code to: $($deployment.Outputs.functionAppName.Value)" -ForegroundColor White
        Write-Host "2. Configure webhook URL in M365: $($deployment.Outputs.apiManagementGatewayUrl.Value)/api/ReceiveEvents" -ForegroundColor White
        Write-Host "3. Get APIM subscription key from Azure portal for webhook authentication" -ForegroundColor White
        Write-Host "4. Test connectivity through the APIM gateway only" -ForegroundColor White

        Write-Host "`n=== Security Features Enabled ===" -ForegroundColor Cyan
        Write-Host "✓ Function App deployed in private VNet" -ForegroundColor Green
        Write-Host "✓ Public access to Function App disabled" -ForegroundColor Green
        Write-Host "✓ APIM configured with VNet integration" -ForegroundColor Green
        Write-Host "✓ NSG rules restrict traffic to M365 service tags" -ForegroundColor Green
        Write-Host "✓ Storage and Key Vault restricted to VNet" -ForegroundColor Green
        Write-Host "✓ All communication over HTTPS with TLS 1.2+" -ForegroundColor Green

    } else {
        Write-Host "Deployment failed with state: $($deployment.ProvisioningState)" -ForegroundColor Red
        exit 1
    }

} catch {
    Write-Host "Error during deployment: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "`nSecure deployment completed successfully!" -ForegroundColor Green
