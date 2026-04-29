param(
    [Parameter(Mandatory=$false, HelpMessage="Specify one or more OUs (Distinguished Names). If omitted, the script defined OUs are used.")]
    [string[]]$SearchBase = @(
        "OU=Name,DC=Name,DC=Name",
        "OU=Name,DC=Name,DC=Name"
    ),
    [switch]$NoGridView
)

# Define the list of attributes (extensionAttribute1 through extensionAttribute15)
$attributes = 1..15 | ForEach-Object { "extensionAttribute$_" }

# Build an LDAP filter: (|(extensionAttribute1=*)(extensionAttribute2=*)...)
# This is more efficient for searching multiple attributes for any value (*)
$ldapFilter = "(|$( ($attributes | ForEach-Object { "($_=*)" }) -join '' ))"

$allUsers = foreach ($ou in $SearchBase) {
    Write-Host "Searching for users in '$ou' with populated extension attributes..." -ForegroundColor Cyan
    try {
        $usersInOu = Get-ADUser -LDAPFilter $ldapFilter -SearchBase $ou -Properties $attributes -ErrorAction Stop
        
        if ($null -ne $usersInOu) {
            Write-Host "Found $($usersInOu.Count) user(s) in '$ou'." -ForegroundColor Green
            $usersInOu
        } else {
            Write-Host "No users found with populated extension attributes in '$ou'." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "An error occurred while querying Active Directory for OU '$ou'. Ensure it is a valid Distinguished Name. Error: $($_.Exception.Message)"
    }
}

if ($allUsers.Count -eq 0) {
    Write-Host "No users found with populated extension attributes across all specified OUs." -ForegroundColor Yellow
} else {
    Write-Host "Found a total of $($allUsers.Count) user(s) across all OUs." -ForegroundColor Green 

    # Transform AD objects into clean PSCustomObjects to avoid "Method Not Supported" errors
    # and ensure all 15 columns are explicitly created for the GridView.
    $report = foreach ($user in $allUsers) {
        $obj = [ordered]@{
            Name              = $user.Name
            UserPrincipalName = $user.UserPrincipalName
        }
        foreach ($attr in $attributes) {
            $obj[$attr] = $user.$attr
        }
        [PSCustomObject]$obj
    }

    if (-not $NoGridView) {
        $report | Out-GridView -Title "Users with Extension Attributes"
    }
    $dateTime = Get-Date -Format "yyyyMMdd-HHmmss"
    $report | Export-Csv -Path ".\Temp\UsersWithExtAttribs-$dateTime.csv" -NoTypeInformation -Encoding UTF8
}

Write-Host "Search complete." -ForegroundColor Cyan