# Deployment script for Gamification Copilot - Sample Tenant
# This script calls the deploy-function.ps1 with tenant-specific parameters

cd ..

.\deploy-function.ps1 `
    -ResourceGroupName "rg-game-inclusion" `
    -Location "uksouth" `
    -ApplicationName "copilogrove2" `
    -EncryptionKey "+Bc" `
    -TenantId "75e67881-b174-484b-9d30-c581c7ebc177" `
    -AuthGuid "0623ba67-0cf2-426a-8f35-8c2670468eeb" `
    -SPO_SiteId "08bfe0b3-c140-46c6-8be3-4bc1133d1f5e" `
    -SPO_ListId "8c07b6ad-74f9-4a8f-8a2d-beeb763842fd" `
    -UserObjectId "e4238485-1d04-4afd-ad31-ea8cab673d94" `
    -IsEmailListExclusive $true `
    -QueueName "user-aggregations" `
    -CopilotEventAggregationsQueueName "copilot-event-aggregations" `
    -EnableTestData $true