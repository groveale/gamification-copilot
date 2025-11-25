
#parameters
$managedIdentityObjectId =""   # 👈 Must be the OBJECT ID of the service principal for the azure function
$deployedFunctionName =""    # 👈 The name of the deployed function ap
$SPO_SiteId =""          # 👈 The SharePoint SiteId that contains the list of exclusion users

# Connect to Graph
    try {
        # Connect to Microsoft Graph - as a global admin or with sufficient permissions
        Write-Step "Connecting to Microsoft Graph..."
        Connect-MgGraph -NoWelcome -Scopes `
            "Sites.FullControl.All", `
            "Application.Read.All", `
            "AppRoleAssignment.ReadWrite.All"

        # Variables (make sure these are set before running the script)
        $principalId = $managedIdentityObjectId     

        # Get clientId from principalId 👈 Used for SharePoint site grant
        $sp = Get-MgServicePrincipal -ServicePrincipalId $principalId
        $clientId = $sp.AppId
      
        $displayName = $deployedFunctionName
        $siteId = $SPO_SiteId                         

        $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $principalId

        # Helper function to check if role is already assigned
        function Test-RoleAssigned($roleId, $resourceId, $assignments) {
            return $assignments | Where-Object {
                $_.AppRoleId -eq $roleId -and $_.ResourceId -eq $resourceId
            }
        }

        ### 1. Assign ActivityFeed.Read from Office 365 Management APIs ###
        $managementAppId = "c5393580-f805-4401-95e8-94b7a6ef2fc2"
        $managementSp = Get-MgServicePrincipal -Filter "appId eq '$managementAppId'"
        $activityFeedRole = $managementSp.AppRoles | Where-Object {
            $_.Value -eq "ActivityFeed.Read" -and $_.AllowedMemberTypes -contains "Application"
        }

        if ($activityFeedRole -and -not (Test-RoleAssigned $activityFeedRole.Id $managementSp.Id $existingAssignments)) {
            $newRole = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $principalId `
                -PrincipalId $principalId `
                -ResourceId $managementSp.Id `
                -AppRoleId $activityFeedRole.Id
        }

        ### 2. Assign Sites.Selected from Microsoft Graph ###
        $graphAppId = "00000003-0000-0000-c000-000000000000"
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
        $sitesSelectedRole = $graphSp.AppRoles | Where-Object {
            $_.Value -eq "Sites.Selected" -and $_.AllowedMemberTypes -contains "Application"
        }
        if ($sitesSelectedRole -and -not (Test-RoleAssigned $sitesSelectedRole.Id $graphSp.Id $existingAssignments)) {
            $newRole = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $principalId `
                -PrincipalId $principalId `
                -ResourceId $graphSp.Id `
                -AppRoleId $sitesSelectedRole.Id
        }

        # Assign Reports.Read.All
        $reportsReadAllRole = $graphSp.AppRoles | Where-Object {
            $_.Value -eq "Reports.Read.All" -and $_.AllowedMemberTypes -contains "Application"
        }
        if ($reportsReadAllRole -and -not (Test-RoleAssigned $reportsReadAllRole.Id $graphSp.Id $existingAssignments)) {
            $newRole = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $principalId `
                -PrincipalId $principalId `
                -ResourceId $graphSp.Id `
                -AppRoleId $reportsReadAllRole.Id
        }

        ### 3. Grant SharePoint permission using Sites.Selected model ###
        $permissionBody = @{
            roles               = @("read") 
            grantedToIdentities = @(
                @{
                    application = @{
                        id          = $clientId       # Must be CLIENT ID here, not objectId
                        displayName = $displayName
                    }
                }
            )
        }

        $newSPOPerms = New-MgSitePermission -SiteId $siteId -BodyParameter $permissionBody

        Write-Status "Managed identity permissions assigned successfully!"
    }
    catch {
        Write-Error "❌ Failed to assign permissions to the managed identity: $($_.Exception.Message)"
        exit 1
    }

