<#
.SYNOPSIS
  Cleans key-based deployment settings, configures identity-based host storage,
  assigns RBAC, and deploys a .NET Azure Function via Functions Core Tools.

.PREREQS
  - Azure CLI (az) installed and logged in: az login
  - Azure Functions Core Tools (func) installed
  - .NET SDK installed (for build)
  - You have Contributor on the function app and storage account scopes

.NOTES
  - Flex Consumption does NOT support WEBSITE_RUN_FROM_PACKAGE, so we do NOT set it.
  - The script is idempotent: deleting non-existent settings and creating
    existing role assignments will be ignored.

#>

param(
  [string]$SubscriptionId        = "cb422ab7-e1e5-47d3-a4c1-4c192f132d3f",
  [string]$ResourceGroup         = "copilot-gamifiy-manual",
  [string]$FunctionAppName       = "copilotgamify",                # change if different
  [string]$StorageAccountName    = "copilotgamifiymanua811b",
  [string]$FunctionMiObjectId    = "7e68c790-ac6b-4bc5-8e52-0d7a77238991",  # System/User-assigned MI objectId
  [string]$ProjectPath           = "C:\Users\alexgrover\source\repos\gamification-copilot\api\src"             # path to your .NET Functions project
)

# ----- Basic checks -----
if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
  Write-Error "Azure CLI (az) not found. Install from https://learn.microsoft.com/cli/azure/install-azure-cli"
  exit 1
}
if (-not (Get-Command func -ErrorAction SilentlyContinue)) {
  Write-Error "Azure Functions Core Tools (func) not found. Install from https://learn.microsoft.com/azure/azure-functions/functions-run-local"
  exit 1
}
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
  Write-Error ".NET SDK (dotnet) not found. Install from https://dotnet.microsoft.com/en-us/download"
  exit 1
}

Write-Host "Setting subscription to $SubscriptionId ..."
az account set --subscription $SubscriptionId | Out-Null

# ----- (0) Ensure the Function App has a managed identity (safe to re-run) -----
Write-Host "Ensuring a system-assigned managed identity exists on $FunctionAppName ..."
az functionapp identity assign `
  --resource-group $ResourceGroup `
  --name $FunctionAppName 1>$null

# ----- (1) RBAC for Storage (Blob + Queue data-plane) -----
$storageScope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Storage/storageAccounts/$StorageAccountName"
$roles = @(
  "Storage Blob Data Contributor",
  "Storage Queue Data Contributor"
)
Write-Host "Assigning Storage data-plane roles to the Function App's MI (ObjectId: $FunctionMiObjectId) ..."
foreach ($role in $roles) {
  az role assignment create `
    --assignee-object-id $FunctionMiObjectId `
    --assignee-principal-type ServicePrincipal `
    --role "$role" `
    --scope $storageScope 2>$null 1>$null
}

# ----- (2) Clean any key-based or legacy settings that can trigger SharedKey usage -----
$toDelete = @(
  "DEPLOYMENT_STORAGE_CONNECTION_STRING",      # causes Kudu to use SharedKey for zip staging
  "AzureWebJobsStorage",                        # old key-based host storage
  "AzureWebJobsDashboard",                      # legacy
  "WEBSITE_CONTENTAZUREFILECONNECTIONSTRING",   # Azure Files content mount (not used on Flex)
  "WEBSITE_CONTENTSHARE",                       # paired with the above
  "WEBSITE_RUN_FROM_PACKAGE"                    # unsupported on Flex
)

Write-Host "Removing key-based / unsupported settings (safe if absent): $($toDelete -join ', ')"
az functionapp config appsettings delete `
  --resource-group $ResourceGroup `
  --name $FunctionAppName `
  --setting-names ($toDelete -join ' ') 1>$null

# ----- (3) Add identity-based host storage settings explicitly -----
# Blob/Queue URIs make the intention explicit for the platform and Kudu checks
$identitySettings = @(
  "AzureWebJobsStorage=DefaultEndpointsProtocol=https;AccountName=$StorageAccountName;EndpointSuffix=core.windows.net",
  "AzureWebJobsStorage__accountName=$StorageAccountName",
  "AzureWebJobsStorage__credential=managedidentity",
  "AzureWebJobsStorage__blobServiceUri=https://$StorageAccountName.blob.core.windows.net",
  "AzureWebJobsStorage__queueServiceUri=https://$StorageAccountName.queue.core.windows.net"
  # If your app uses TABLES via AzureWebJobsStorage, also add:
  # "AzureWebJobsStorage__tableServiceUri=https://$StorageAccountName.table.core.windows.net"
)

Write-Host "Applying identity-based host storage settings ..."
az functionapp config appsettings set `
  --resource-group $ResourceGroup `
  --name $FunctionAppName `
  --settings $identitySettings | Out-Null

# ----- (4) Sanity check: show any app settings that still look like SharedKey/SAS -----
Write-Host "`nScanning for any remaining key-based values (should return nothing):"
$cfg = az functionapp config appsettings list -g $ResourceGroup -n $FunctionAppName | ConvertFrom-Json
$leftovers = $cfg | Where-Object { $_.value -match 'AccountKey=|SharedAccessSignature=|DefaultEndpointsProtocol=' }
if ($leftovers) { $leftovers | Format-Table name, value -AutoSize } else { Write-Host "✔ No key-based settings detected." }

# ----- (5) Restart to ensure Kudu/host picks up config -----
Write-Host "`nRestarting Function App ..."
az functionapp restart -g $ResourceGroup -n $FunctionAppName 1>$null

# ----- (6) Build & Deploy (.NET) using Core Tools (AAD token from az login) -----
Write-Host "`nBuilding and publishing the .NET Azure Function using Functions Core Tools ..."
Push-Location $ProjectPath
try {
  dotnet restore
  dotnet build -c Release
  # publish will rebuild if needed; keeping explicit build makes failures clearer
  func azure functionapp publish $FunctionAppName --dotnet-isolated
}
finally {
  Pop-Location
}

Write-Host "`nDone. If deployment still fails with KeyBasedAuthenticationNotPermitted:"
Write-Host " - Re-check there is NO 'DEPLOYMENT_STORAGE_CONNECTION_STRING' on the app or any slot."
Write-Host " - Confirm the MI has Blob/Queue Data Contributor on storage."
Write-Host " - If storage uses Private Endpoints/firewall, ensure the app can reach it (VNet integration)."