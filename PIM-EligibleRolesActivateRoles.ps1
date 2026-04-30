# Make sure you are connected to Microsoft Graph with the necessary scopes.
# For this script, you will need at least "RoleManagement.ReadWrite.Directory" to activate roles.
# Example: Connect-MgGraph -Scopes "User.Read.All", "RoleManagement.ReadWrite.Directory"

# Dynamically get the current user's ID and Name from the Graph connection context
$context = Get-MgContext
$me = Get-MgUser -UserId $context.Account -Property Id, DisplayName -ErrorAction Stop
$currentUser = $me.Id
$currentName = $me.DisplayName

Write-Host "Fetching eligible roles for '$currentName' ($currentUser)..." -ForegroundColor Cyan

# Get all eligible roles for the current user
try {
    $eligibleRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentUser'" -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve eligible roles. Ensure you are connected to Microsoft Graph with 'RoleManagement.Read.Directory' scope and have sufficient permissions. Error: $($_.Exception.Message)"
    exit
}

if (-not $eligibleRoles) {
    Write-Warning "No eligible roles found for '$currentName'."
    exit
}

Write-Host "Displaying eligible roles. Select one or more roles to activate." -ForegroundColor Green

# Prepare roles for Out-GridView display and selection
# We include the RoleDefinitionId and DirectoryScopeId as they are needed for activation,
# but hide them from the default view in Out-GridView by not explicitly selecting them for display
$rolesForSelection = $eligibleRoles | Select-Object @{Name = "RoleDisplayName"; Expression = { $_.RoleDefinition.DisplayName }}, RoleDefinitionId, DirectoryScopeId, Id, CreatedDateTime, MemberType | Sort-Object RoleDisplayName

# Display available roles with index numbers for command-line selection
Write-Host "`nAvailable Roles:" -ForegroundColor Cyan
for ($i = 0; $i -lt $rolesForSelection.Count; $i++) {
    Write-Host "[$($i + 1)] $($rolesForSelection[$i].RoleDisplayName)"
}

$selectionInput = Read-Host "`nEnter the number(s) of the roles to activate (comma-separated, e.g., 1,3) or 'A' for all"

if ($selectionInput -match "^[Aa]$") {
    $selectedRoles = $rolesForSelection
} else {
    $indices = $selectionInput -split ',' | ForEach-Object { 
        $val = $_.Trim()
        if ($val -as [int]) { [int]$val - 1 }
    }
    $validIndices = $indices | Where-Object { $_ -ge 0 -and $_ -lt $rolesForSelection.Count }
    $selectedRoles = $rolesForSelection[$validIndices]
}

if (-not $selectedRoles) {
    Write-Host "No roles selected. Exiting." -ForegroundColor Yellow
    exit
}

Write-Host "`nSelected Roles for Activation:" -ForegroundColor Magenta
$selectedRoles | Format-Table -Property RoleDisplayName, CreatedDateTime, MemberType -AutoSize

# --- Prompt for Activation Details ---
$justification = Read-Host "Enter a justification for activating these roles"
if ([string]::IsNullOrWhiteSpace($justification)) {
    Write-Warning "Justification cannot be empty. Exiting."
    exit
}

$durationHours = [int](Read-Host "Enter the duration in hours for the activation (e.g., 1, 4, 8) - default 8 hours")
if ($durationHours -le 0) {
    Write-Host "Invalid duration. Defaulting to 8 hours." -ForegroundColor Yellow
    $durationHours = 8
}

# --- Activate each selected role ---
foreach ($roleToActivate in $selectedRoles) {
    Write-Host "`nAttempting to activate role: $($roleToActivate.RoleDisplayName)" -ForegroundColor Cyan

    $params = @{
        Action           = "selfActivate"
        PrincipalId      = $currentUser # The user activating the role is the principal
        RoleDefinitionId = $roleToActivate.RoleDefinitionId
        DirectoryScopeId = $roleToActivate.DirectoryScopeId
        Justification    = $justification
        ScheduleInfo = @{
            StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") # PIM expects UTC and specific ISO 8601 format
            Expiration    = @{
                Type     = "AfterDuration"
                Duration = "PT$($durationHours)H" # PIM duration format: PTnH for n hours
            }
        }
    }

    try {
        $activationRequest = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params -ErrorAction Stop
        Write-Host "Successfully submitted activation request for $($roleToActivate.RoleDisplayName)." -ForegroundColor Green
        Write-Host "Request ID: $($activationRequest.Id)" -ForegroundColor Green
        Write-Host "Request Status: $($activationRequest.Status)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to activate role $($roleToActivate.RoleDisplayName). Error: $($_.Exception.Message)"
    }
}

Write-Host "`nActivation process complete." -ForegroundColor DarkGreen