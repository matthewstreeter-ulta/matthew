# Prerequisites:
# Install-Module Microsoft.Graph.Groups
# Install-Module Microsoft.Graph.Users.Authentication

#region Configuration
$SecurityGroupName = "Persona_Admin" # <--- IMPORTANT: Replace with the actual name of your security group
$OutputCsvPath     = ".\Temp\UsersWithoutAuthenticator-$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
#endregion

#region Connect to Microsoft Graph
# Scopes required:
# Group.Read.All - To read group memberships
# User.Read.All - To read basic user profiles
# UserAuthenticationMethod.Read.All - To read user authentication methods
# AuditLog.Read.All - To read sign-in activity (for SignInActivity property)
$RequiredScopes = @("Group.Read.All", "User.Read.All", "UserAuthenticationMethod.Read.All", "AuditLog.Read.All")

Write-Host "Connecting to Microsoft Graph with required scopes..." -ForegroundColor Cyan
try {
    # Attempt to connect. If already connected with sufficient scopes, it will reuse the connection.
    # Otherwise, it will prompt for authentication.
    Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Please ensure you have the necessary permissions and try again. Error: $($_.Exception.Message)"
    return
}
#endregion

#region Get Security Group and Members
Write-Host "Searching for security group: '$SecurityGroupName'..." -ForegroundColor Cyan
try {
    $group = Get-MgGroup -Filter "DisplayName eq '$SecurityGroupName'" -ErrorAction Stop
    if (-not $group) {
        Write-Warning "No security group found with the name '$SecurityGroupName'."
        Disconnect-MgGraph
        return
    }
    Write-Host "Found group: $($group.DisplayName) (ID: $($group.Id))" -ForegroundColor Green
}
catch {
    Write-Error "Error finding security group: $($_.Exception.Message)"
    Disconnect-MgGraph
    return
}

Write-Host "Retrieving members of '$($group.DisplayName)'..." -ForegroundColor Cyan
$groupMembers = @()
try {
    # -All parameter handles pagination
    $groupMembers = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }
    Write-Host "Found $($groupMembers.Count) user members in the group." -ForegroundColor Green
}
catch {
    Write-Error "Error retrieving group members: $($_.Exception.Message)"
    Disconnect-MgGraph
    return
}

if ($groupMembers.Count -eq 0) {
    Write-Warning "No user members found in the security group. Exiting."
    Disconnect-MgGraph
    return
}
#endregion

#region Check Authentication Methods for Each User
$reportResults = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Checking authentication methods for each user..." -ForegroundColor Cyan
foreach ($member in $groupMembers) {
    $userId = $member.Id

    # Fetch user details including SignInActivity
    try {
        $fullUser = Get-MgUser -UserId $userId -Property UserPrincipalName, DisplayName, SignInActivity -ErrorAction Stop
        $userPrincipalName = $fullUser.UserPrincipalName
        $displayName       = $fullUser.DisplayName
    }
    catch {
        Write-Warning "Could not retrieve details for user ID: $userId. Skipping."
        continue
    }

    # Fetch Manager info
    $managerName = "N/A"
    $managerMail = "N/A"
    try {
        $manager = Get-MgUserManager -UserId $userId -ErrorAction SilentlyContinue
        if ($null -ne $manager) {
            $managerName = $manager.AdditionalProperties['displayName']
            $managerMail = $manager.AdditionalProperties['mail']
            if ([string]::IsNullOrWhiteSpace($managerMail)) { $managerMail = $manager.AdditionalProperties['userPrincipalName'] }
        }
    } catch {
        # Manager might not be assigned or is inaccessible
    }

    Write-Host "Processing user: $displayName ($userPrincipalName)" -NoNewline

    $hasAuthenticator = $false
    try {
        $authMethods = Get-MgUserAuthenticationMethod -UserId $userId -ErrorAction SilentlyContinue
        foreach ($method in $authMethods) {
            # The Microsoft Authenticator app typically registers as a MicrosoftAuthenticatorAuthenticationMethod
            # or sometimes as a Fido2AuthenticationMethod if it's a passwordless sign-in.
            # For this specific request, we'll focus on MicrosoftAuthenticatorAuthenticationMethod.
            if ($method.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod') {
                $hasAuthenticator = $true
                break
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve authentication methods for user $userPrincipalName. Error: $($_.Exception.Message)"
        # Assume no authenticator if we can't check
        $hasAuthenticator = $false
    }

    $status = if ($hasAuthenticator) { "Registered" } else { "Not Registered" }
    if ($hasAuthenticator) {
        Write-Host " -> Microsoft Authenticator registered." -ForegroundColor Green
    }
    else {
        Write-Host " -> NO Microsoft Authenticator registered." -ForegroundColor Red
    }

    $reportResults.Add([PSCustomObject]@{
        DisplayName                  = $displayName
        UserPrincipalName            = $userPrincipalName
        UserId                       = $userId
        ManagerName                  = $managerName
        ManagerEmail                 = $managerMail
        AuthenticatorStatus          = $status
        LastSignInDateTime           = $fullUser.SignInActivity.LastSignInDateTime
        LastSuccessfulSignInDateTime = $fullUser.SignInActivity.LastSuccessfulSignInDateTime
        Group                        = $group.DisplayName
    })
}
#endregion

#region Output Results
Write-Host "`nProcessing complete. Exporting results for $($reportResults.Count) users..." -ForegroundColor Cyan
$reportResults | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Details exported to: $OutputCsvPath" -ForegroundColor Green

Disconnect-MgGraph
Write-Host "Script finished." -ForegroundColor Cyan
#endregion
