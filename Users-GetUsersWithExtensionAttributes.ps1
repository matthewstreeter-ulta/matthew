param(
    [Parameter(Mandatory=$false, HelpMessage="The number of the extensionAttribute to check (1-15). Set to 0 or omit to check all 1-15.")]
    [ValidateRange(0, 15)]
    [int]$AttributeNumber = 0,

    [Parameter(Mandatory=$false, HelpMessage="The filter mode: 'Missing' (attribute not set) or 'Present' (attribute has a value).")]
    [ValidateSet("Missing", "Present")]
    [string]$FilterMode = "Missing",

    [Parameter(Mandatory=$false, HelpMessage="Specify one or more OUs (Distinguished Names). If omitted, the script defined OUs are used.")]
    [string[]]$SearchBase = @(
        "OU=Dept,DC=domain,DC=com", #Replace with your OUs
        "OU=Dept,DC=domain,DC=com"
    ),
    [switch]$NoGridView
)

# Define the list of attributes (extensionAttribute1 through extensionAttribute15)
$attributes = 1..15 | ForEach-Object { "extensionAttribute$_" }

# Build an LDAP filter based on parameters
if ($AttributeNumber -ge 1 -and $AttributeNumber -le 15) {
    $targetAttribute = "extensionAttribute$AttributeNumber"
    $attributeFilter = if ($FilterMode -eq "Missing") { "(!($targetAttribute=*))" } else { "($targetAttribute=*)" }
}
else {
    # Default behavior: Check across all 15 attributes
    $targetAttribute = "extensionAttributes1-15"
    $operator = if ($FilterMode -eq "Missing") { "&" } else { "|" }
    $subFilters = 1..15 | ForEach-Object { 
        if ($FilterMode -eq "Missing") { "(!(extensionAttribute$_=*))" } else { "(extensionAttribute$_=*)" }
    }
    $attributeFilter = "($operator$($subFilters -join ''))"
}

$ldapFilter = "(&$attributeFilter(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" # Exclude disabled users

$allUsers = foreach ($ou in $SearchBase) {
    Write-Host "Searching for enabled users in '$ou' where $targetAttribute is $FilterMode..." -ForegroundColor Cyan
    try {
        $usersInOu = Get-ADUser -LDAPFilter $ldapFilter -SearchBase $ou -Properties ($attributes + "displayName") -ErrorAction Stop
        
        if ($null -ne $usersInOu) {
            Write-Host "Found $($usersInOu.Count) user(s) in '$ou'." -ForegroundColor Green
            $usersInOu
        } else {
            Write-Host "No enabled users found where $targetAttribute is $FilterMode in '$ou'." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "An error occurred while querying Active Directory for OU '$ou'. Ensure it is a valid Distinguished Name. Error: $($_.Exception.Message)"
    }
}

if ($allUsers.Count -eq 0) {
    Write-Host "No enabled users found where $targetAttribute is $FilterMode across all specified OUs." -ForegroundColor Yellow
} else {
    Write-Host "Found a total of $($allUsers.Count) user(s) across all OUs." -ForegroundColor Green 

    # Transform AD objects into clean PSCustomObjects to avoid "Method Not Supported" errors
    # and ensure all 15 columns are explicitly created for the GridView.
    $report = foreach ($user in $allUsers) {
        $obj = [ordered]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            # Extract the immediate parent OU name from the DistinguishedName
            OUName            = ($user.DistinguishedName -split ',') |
                                Where-Object { $_ -like 'OU=*' } |
                                Select-Object -First 1 |
                                ForEach-Object { $_.Substring(3) }
        }
        # Handle cases where a user might be directly in the domain root (no OU)
        if ([string]::IsNullOrEmpty($obj.OUName)) {
            $obj.OUName = "Domain Root"
        }
        foreach ($attr in $attributes) {
            $obj[$attr] = $user.$attr
        }
        [PSCustomObject]$obj
    }

    if (-not $NoGridView) {
        $report | Out-GridView -Title "Enabled Users - $targetAttribute is $FilterMode"
    }
    $dateTime = Get-Date -Format "yyyyMMdd-HHmmss"
    $report | Export-Csv -Path ".\Temp\EnabledUsers-$targetAttribute-$FilterMode-$dateTime.csv" -NoTypeInformation -Encoding UTF8
}

Write-Host "Search complete." -ForegroundColor Cyan
