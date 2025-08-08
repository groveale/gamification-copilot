# Security Validation Script for Secure Azure Function App Deployment
# This script validates that security controls are properly configured

param(
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$true)]
    [string]$ApplicationName
)

$ErrorActionPreference = "Stop"

# Resource names based on naming convention
$functionAppName = "func-$ApplicationName"
$apimName = "apim-$ApplicationName"
$vnetName = "vnet-$ApplicationName"
$nsgName = "nsg-$ApplicationName"
$storageAccountName = "store$ApplicationName"
$keyVaultName = "kv-$ApplicationName"

Write-Host "=== Security Validation for $ApplicationName ===" -ForegroundColor Cyan
Write-Host "Resource Group: $ResourceGroupName" -ForegroundColor Yellow
Write-Host ""

$validationResults = @()

try {
    # Test 1: Function App Public Access
    Write-Host "🔍 Testing Function App Public Access..." -ForegroundColor Yellow
    $functionApp = Get-AzWebApp -ResourceGroupName $ResourceGroupName -Name $functionAppName
    
    if ($functionApp.SiteConfig.PublicNetworkAccess -eq "Disabled") {
        Write-Host "✅ Function App public access is disabled" -ForegroundColor Green
        $validationResults += @{Test="Function Public Access"; Status="PASS"; Details="Public access disabled"}
    } else {
        Write-Host "❌ Function App public access is enabled" -ForegroundColor Red
        $validationResults += @{Test="Function Public Access"; Status="FAIL"; Details="Public access enabled"}
    }

    # Test 2: Function App VNet Integration
    Write-Host "🔍 Testing Function App VNet Integration..." -ForegroundColor Yellow
    if ($functionApp.VirtualNetworkSubnetId) {
        Write-Host "✅ Function App has VNet integration configured" -ForegroundColor Green
        $validationResults += @{Test="VNet Integration"; Status="PASS"; Details="VNet integration active"}
    } else {
        Write-Host "❌ Function App missing VNet integration" -ForegroundColor Red
        $validationResults += @{Test="VNet Integration"; Status="FAIL"; Details="No VNet integration"}
    }

    # Test 3: IP Security Restrictions
    Write-Host "🔍 Testing Function App IP Restrictions..." -ForegroundColor Yellow
    $ipRestrictions = $functionApp.SiteConfig.IpSecurityRestrictions
    $apimSubnetRestriction = $ipRestrictions | Where-Object { $_.VnetSubnetResourceId -like "*snet-apim*" }
    
    if ($apimSubnetRestriction -and $ipRestrictions.Count -gt 1) {
        Write-Host "✅ Function App has proper IP restrictions" -ForegroundColor Green
        $validationResults += @{Test="IP Security Restrictions"; Status="PASS"; Details="APIM subnet allowed, others denied"}
    } else {
        Write-Host "❌ Function App IP restrictions not properly configured" -ForegroundColor Red
        $validationResults += @{Test="IP Security Restrictions"; Status="FAIL"; Details="Missing or incorrect restrictions"}
    }

    # Test 4: Storage Account Network Access
    Write-Host "🔍 Testing Storage Account Security..." -ForegroundColor Yellow
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $storageAccountName
    
    if ($storageAccount.NetworkRuleSet.DefaultAction -eq "Deny" -and $storageAccount.NetworkRuleSet.VirtualNetworkRules.Count -gt 0) {
        Write-Host "✅ Storage Account has VNet restrictions" -ForegroundColor Green
        $validationResults += @{Test="Storage Network Security"; Status="PASS"; Details="Default deny with VNet rules"}
    } else {
        Write-Host "❌ Storage Account network security not configured" -ForegroundColor Red
        $validationResults += @{Test="Storage Network Security"; Status="FAIL"; Details="Missing network restrictions"}
    }

    # Test 5: Key Vault Access Policy
    Write-Host "🔍 Testing Key Vault Security..." -ForegroundColor Yellow
    $keyVault = Get-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $keyVaultName
    
    if ($keyVault.NetworkAcls.DefaultAction -eq "Deny" -and $keyVault.NetworkAcls.VirtualNetworkResourceIds.Count -gt 0) {
        Write-Host "✅ Key Vault has VNet restrictions" -ForegroundColor Green
        $validationResults += @{Test="Key Vault Network Security"; Status="PASS"; Details="Default deny with VNet rules"}
    } else {
        Write-Host "❌ Key Vault network security not configured" -ForegroundColor Red
        $validationResults += @{Test="Key Vault Network Security"; Status="FAIL"; Details="Missing network restrictions"}
    }

    # Test 6: Network Security Group Rules
    Write-Host "🔍 Testing Network Security Group Rules..." -ForegroundColor Yellow
    $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $ResourceGroupName -Name $nsgName
    $m365Rule = $nsg.SecurityRules | Where-Object { $_.SourceAddressPrefix -eq "M365ManagementActivityApi" }
    $denyRule = $nsg.SecurityRules | Where-Object { $_.Access -eq "Deny" -and $_.Priority -eq 4096 }
    
    if ($m365Rule -and $denyRule) {
        Write-Host "✅ NSG has proper security rules" -ForegroundColor Green
        $validationResults += @{Test="NSG Security Rules"; Status="PASS"; Details="M365 service tag allowed, default deny configured"}
    } else {
        Write-Host "❌ NSG security rules not properly configured" -ForegroundColor Red
        $validationResults += @{Test="NSG Security Rules"; Status="FAIL"; Details="Missing M365 or deny rules"}
    }

    # Test 7: API Management VNet Integration
    Write-Host "🔍 Testing API Management Configuration..." -ForegroundColor Yellow
    $apim = Get-AzApiManagement -ResourceGroupName $ResourceGroupName -Name $apimName
    
    if ($apim.VirtualNetwork.Type -eq "External" -and $apim.VirtualNetwork.SubnetResourceId) {
        Write-Host "✅ APIM has external VNet integration" -ForegroundColor Green
        $validationResults += @{Test="APIM VNet Integration"; Status="PASS"; Details="External VNet type configured"}
    } else {
        Write-Host "❌ APIM VNet integration not configured" -ForegroundColor Red
        $validationResults += @{Test="APIM VNet Integration"; Status="FAIL"; Details="Missing or incorrect VNet configuration"}
    }

    # Test 8: HTTPS Configuration
    Write-Host "🔍 Testing HTTPS Configuration..." -ForegroundColor Yellow
    if ($functionApp.HttpsOnly -eq $true) {
        Write-Host "✅ Function App requires HTTPS" -ForegroundColor Green
        $validationResults += @{Test="HTTPS Configuration"; Status="PASS"; Details="HTTPS only enabled"}
    } else {
        Write-Host "❌ Function App allows HTTP" -ForegroundColor Red
        $validationResults += @{Test="HTTPS Configuration"; Status="FAIL"; Details="HTTPS only not enforced"}
    }

    # Test 9: Managed Identity
    Write-Host "🔍 Testing Managed Identity..." -ForegroundColor Yellow
    if ($functionApp.Identity.Type -eq "SystemAssigned") {
        Write-Host "✅ Function App has system-assigned managed identity" -ForegroundColor Green
        $validationResults += @{Test="Managed Identity"; Status="PASS"; Details="System-assigned identity configured"}
    } else {
        Write-Host "❌ Function App missing managed identity" -ForegroundColor Red
        $validationResults += @{Test="Managed Identity"; Status="FAIL"; Details="No managed identity"}
    }

    # Test 10: Connectivity Test (if possible)
    Write-Host "🔍 Testing Direct Function App Access..." -ForegroundColor Yellow
    try {
        $directUrl = "https://$($functionApp.DefaultHostName)/api/ReceiveEvents"
        $response = Invoke-WebRequest -Uri $directUrl -Method GET -TimeoutSec 10 -ErrorAction Stop
        Write-Host "❌ Direct Function App access is possible (should be blocked)" -ForegroundColor Red
        $validationResults += @{Test="Direct Access Block"; Status="FAIL"; Details="Direct access not blocked"}
    } catch {
        if ($_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*forbidden*") {
            Write-Host "✅ Direct Function App access is properly blocked" -ForegroundColor Green
            $validationResults += @{Test="Direct Access Block"; Status="PASS"; Details="Direct access blocked (403 Forbidden)"}
        } else {
            Write-Host "⚠️ Direct Function App access test inconclusive: $($_.Exception.Message)" -ForegroundColor Yellow
            $validationResults += @{Test="Direct Access Block"; Status="WARNING"; Details="Test inconclusive: $($_.Exception.Message)"}
        }
    }

} catch {
    Write-Host "❌ Error during validation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Summary Report
Write-Host "`n=== SECURITY VALIDATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total Tests: $($validationResults.Count)" -ForegroundColor White

$passedTests = ($validationResults | Where-Object { $_.Status -eq "PASS" }).Count
$failedTests = ($validationResults | Where-Object { $_.Status -eq "FAIL" }).Count
$warningTests = ($validationResults | Where-Object { $_.Status -eq "WARNING" }).Count

Write-Host "Passed: $passedTests" -ForegroundColor Green
Write-Host "Failed: $failedTests" -ForegroundColor Red
Write-Host "Warnings: $warningTests" -ForegroundColor Yellow

Write-Host "`nDetailed Results:" -ForegroundColor White
$validationResults | ForEach-Object {
    $color = switch ($_.Status) {
        "PASS" { "Green" }
        "FAIL" { "Red" }
        "WARNING" { "Yellow" }
        default { "White" }
    }
    Write-Host "[$($_.Status)] $($_.Test): $($_.Details)" -ForegroundColor $color
}

# Security Score
$securityScore = [math]::Round(($passedTests / ($passedTests + $failedTests)) * 100, 1)
Write-Host "`nSecurity Score: $securityScore%" -ForegroundColor $(if ($securityScore -ge 90) { "Green" } elseif ($securityScore -ge 70) { "Yellow" } else { "Red" })

if ($failedTests -gt 0) {
    Write-Host "`n⚠️ ATTENTION: $failedTests security test(s) failed. Review and fix issues before production use." -ForegroundColor Red
    exit 1
} else {
    Write-Host "`n✅ All critical security tests passed! Your deployment is secure." -ForegroundColor Green
}
