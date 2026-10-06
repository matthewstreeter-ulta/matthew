<#
.SYNOPSIS
    Exports all Microsoft Entra Conditional Access policies and resolves
    referenced object IDs into human-readable names.

.DESCRIPTION
    This all-in-one version performs the entire workflow:

        1. Connect to Microsoft Graph
        2. Export all Conditional Access policies as individual JSON files
        3. Preserve those original JSON exports
        4. Resolve object IDs referenced by each policy
        5. Export human-readable resolved JSON copies
        6. Create a detailed assignment CSV
        7. Create a one-row-per-policy summary CSV

    Resolves:
        - Users
        - Groups
        - Applications / Cloud Apps
        - Service Principals
        - Directory Roles
        - Named Locations

    Original exported JSON files are NOT modified.

.REQUIREMENTS
    Microsoft.Graph.Authentication
    Microsoft.Graph.Identity.SignIns

    Delegated Graph permissions:
        Policy.Read.All
        Directory.Read.All
        Application.Read.All

.EXAMPLE
    .\CA-ExportPolicies_ResolveGuids.ps1.ps1 -OutputFolder "C:\Temp\CAPolicies"

.NOTES
    Conditions.Applications.IncludeApplications and
    Conditions.Applications.ExcludeApplications generally contain
    application AppId / Client ID values.

    Conditions.ClientApplications.IncludeServicePrincipals and
    ExcludeServicePrincipals contain Service Principal Object IDs.
#>

[CmdletBinding()]
param (

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\Temp\CAPolicies"

)

# ============================================================
# Output paths
# ============================================================

$OriginalJsonFolder = Join-Path `
    $OutputFolder `
    "Original"

$ResolvedJsonFolder = Join-Path `
    $OutputFolder `
    "Resolved"

$DetailedCsvPath = Join-Path `
    $OutputFolder `
    "ConditionalAccessPolicy-ResolvedAssignments.csv"

$PolicySummaryCsvPath = Join-Path `
    $OutputFolder `
    "ConditionalAccessPolicy-PolicySummary.csv"

$null = New-Item `
    -ItemType Directory `
    -Path $OutputFolder `
    -Force

$null = New-Item `
    -ItemType Directory `
    -Path $OriginalJsonFolder `
    -Force

$null = New-Item `
    -ItemType Directory `
    -Path $ResolvedJsonFolder `
    -Force

# ============================================================
# Validate Graph modules
# ============================================================

$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Identity.SignIns"
)

foreach ($Module in $RequiredModules) {

    if (-not (Get-Module -ListAvailable $Module)) {

        throw @"
Required module is not installed:

    $Module

Install it with:

    Install-Module $Module -Scope CurrentUser

"@
    }
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.SignIns

# ============================================================
# Connect to Microsoft Graph
# ============================================================

$RequiredScopes = @(
    "Policy.Read.All",
    "Directory.Read.All",
    "Application.Read.All"
)

$GraphContext = Get-MgContext

if (-not $GraphContext) {

    Write-Host ""
    Write-Host "Connecting to Microsoft Graph..." `
        -ForegroundColor Cyan

    Connect-MgGraph `
        -Scopes $RequiredScopes `
        -NoWelcome

    $GraphContext = Get-MgContext
}
else {

    Write-Host ""
    Write-Host "Existing Microsoft Graph connection detected." `
        -ForegroundColor DarkGray

    Write-Host "Account : $($GraphContext.Account)"
    Write-Host "Tenant  : $($GraphContext.TenantId)"
}

# ============================================================
# Runtime statistics
# ============================================================

$Stats = [ordered]@{
    PoliciesExported         = 0
    PoliciesProcessed        = 0
    PoliciesFailed           = 0

    UserLookups              = 0
    GroupLookups             = 0
    ApplicationLookups       = 0
    ServicePrincipalLookups  = 0
    RoleLookups              = 0
    LocationLookups          = 0

    CacheHits                = 0
    ResolutionFailures       = 0
}

$TotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ============================================================
# Export Conditional Access policies
# ============================================================

Write-Host ""
Write-Host "Exporting Conditional Access Policies" `
    -ForegroundColor Cyan
Write-Host "=====================================" `
    -ForegroundColor Cyan
Write-Host ""

try {

    $Policies = @(
        Get-MgIdentityConditionalAccessPolicy `
            -All `
            -ErrorAction Stop
    )
}
catch {

    throw @"
Unable to retrieve Conditional Access policies.

$($_.Exception.Message)
"@
}

if ($Policies.Count -eq 0) {

    throw "No Conditional Access policies were returned from Microsoft Graph."
}

Write-Host "Policies discovered: $($Policies.Count)"
Write-Host ""

$ExportNumber = 0

foreach ($Policy in $Policies) {

    $ExportNumber++

    Write-Progress `
        -Activity "Exporting Conditional Access Policies" `
        -Status (
            "Policy {0} of {1}: {2}" -f
            $ExportNumber,
            $Policies.Count,
            $Policy.DisplayName
        ) `
        -PercentComplete (
            ($ExportNumber / $Policies.Count) * 100
        )

    $SafeName = (
        $Policy.DisplayName -replace '[\\/:*?"<>|]', '_'
    )

    # Avoid filename collisions if two policies have the same name.
    $ExportPath = Join-Path `
        $OriginalJsonFolder `
        "$SafeName.json"

    if (Test-Path $ExportPath) {

        $ExportPath = Join-Path `
            $OriginalJsonFolder `
            "$SafeName-$($Policy.Id).json"
    }

    $Policy |
        ConvertTo-Json -Depth 100 |
        Set-Content `
            -Path $ExportPath `
            -Encoding UTF8

    $Stats.PoliciesExported++
}

Write-Progress `
    -Activity "Exporting Conditional Access Policies" `
    -Completed

Write-Host "Export complete." `
    -ForegroundColor Green

Write-Host "Original JSON folder:"
Write-Host "  $OriginalJsonFolder"
Write-Host ""

# ============================================================
# Lookup caches
# ============================================================

$UserCache             = @{}
$GroupCache            = @{}
$ApplicationCache      = @{}
$ServicePrincipalCache = @{}
$RoleCache             = @{}
$LocationCache         = @{}

# ============================================================
# Reports
# ============================================================

$ReportRows =
    [System.Collections.Generic.List[object]]::new()

$PolicySummaryRows =
    [System.Collections.Generic.List[object]]::new()

# ============================================================
# Built-in CA values
# ============================================================

$SpecialValues = @{
    "All"                   = "All"
    "None"                  = "None"
    "GuestsOrExternalUsers" = "Guests or External Users"
    "Office365"             = "Office 365"
    "MicrosoftAdminPortals" = "Microsoft Admin Portals"
    "AllTrusted"            = "All Trusted Locations"
    "AllCompliant"          = "All Compliant"
    "AllDomainJoined"       = "All Domain Joined"
}

# ============================================================
# Helpers
# ============================================================

function ConvertTo-UrlEncodedValue {

    param (
        [Parameter(Mandatory)]
        [string]$Value
    )

    return [System.Uri]::EscapeDataString($Value)
}

function Get-SpecialCAValue {

    param (
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    if ($SpecialValues.ContainsKey($Value)) {

        return [PSCustomObject]@{
            Id                = $Value
            AppId             = $null
            DisplayName       = $SpecialValues[$Value]
            UserPrincipalName = $null
            Mail              = $null
            ObjectType        = "Special"
            Status            = "BuiltInValue"
            ResolutionMessage = $null
        }
    }

    return $null
}

function New-UnresolvedResult {

    param (
        [string]$Id,
        [string]$AppId,
        [string]$ObjectType,
        [string]$Message
    )

    return [PSCustomObject]@{
        Id                = $Id
        AppId             = $AppId
        DisplayName       = $null
        UserPrincipalName = $null
        Mail              = $null
        ObjectType        = $ObjectType
        Status            = "PossiblyDeletedOrUnavailable"
        ResolutionMessage = $Message
    }
}

# ============================================================
# User resolver
# ============================================================

function Resolve-CAUser {

    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    $Special = Get-SpecialCAValue -Value $Id

    if ($Special) {
        return $Special
    }

    if ($UserCache.ContainsKey($Id)) {

        $Stats.CacheHits++

        return $UserCache[$Id]
    }

    $Stats.UserLookups++

    try {

        $Uri =
            "https://graph.microsoft.com/v1.0/users/$Id" +
            '?$select=id,displayName,userPrincipalName,mail,userType,accountEnabled'

        $User = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $Result = [PSCustomObject]@{
            Id                = $User.id
            AppId             = $null
            DisplayName       = $User.displayName
            UserPrincipalName = $User.userPrincipalName
            Mail              = $User.mail
            UserType          = $User.userType
            AccountEnabled    = $User.accountEnabled
            ObjectType        = "User"
            Status            = "Resolved"
            ResolutionMessage = $null
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $Id `
            -AppId $null `
            -ObjectType "User" `
            -Message $_.Exception.Message
    }

    $UserCache[$Id] = $Result

    return $Result
}

# ============================================================
# Group resolver
# ============================================================

function Resolve-CAGroup {

    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ($GroupCache.ContainsKey($Id)) {

        $Stats.CacheHits++

        return $GroupCache[$Id]
    }

    $Stats.GroupLookups++

    try {

        $Uri =
            "https://graph.microsoft.com/v1.0/groups/$Id" +
            '?$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes'

        $Group = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $Result = [PSCustomObject]@{
            Id                = $Group.id
            AppId             = $null
            DisplayName       = $Group.displayName
            UserPrincipalName = $null
            Mail              = $Group.mail
            SecurityEnabled   = $Group.securityEnabled
            MailEnabled       = $Group.mailEnabled
            GroupTypes        = @($Group.groupTypes)
            ObjectType        = "Group"
            Status            = "Resolved"
            ResolutionMessage = $null
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $Id `
            -AppId $null `
            -ObjectType "Group" `
            -Message $_.Exception.Message
    }

    $GroupCache[$Id] = $Result

    return $Result
}

# ============================================================
# Application resolver
# ============================================================

function Resolve-CAApplication {

    param (
        [Parameter(Mandatory)]
        [string]$AppId
    )

    $Special = Get-SpecialCAValue -Value $AppId

    if ($Special) {
        return $Special
    }

    if ($ApplicationCache.ContainsKey($AppId)) {

        $Stats.CacheHits++

        return $ApplicationCache[$AppId]
    }

    $Stats.ApplicationLookups++

    try {

        $Filter = ConvertTo-UrlEncodedValue `
            -Value "appId eq '$AppId'"

        $Uri =
            "https://graph.microsoft.com/v1.0/servicePrincipals" +
            "?`$filter=$Filter" +
            "&`$select=id,appId,displayName,servicePrincipalType,accountEnabled"

        $Response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $SP = @($Response.value) |
            Select-Object -First 1

        if ($SP) {

            $Result = [PSCustomObject]@{
                Id                   = $SP.id
                AppId                = $SP.appId
                DisplayName          = $SP.displayName
                UserPrincipalName    = $null
                Mail                 = $null
                ServicePrincipalType = $SP.servicePrincipalType
                AccountEnabled       = $SP.accountEnabled
                ObjectType           = "Application"
                Status               = "Resolved"
                ResolutionMessage    = $null
            }
        }
        else {

            $Uri =
                "https://graph.microsoft.com/v1.0/applications" +
                "?`$filter=$Filter" +
                "&`$select=id,appId,displayName"

            $Response = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $Uri `
                -ErrorAction Stop

            $App = @($Response.value) |
                Select-Object -First 1

            if ($App) {

                $Result = [PSCustomObject]@{
                    Id                   = $App.id
                    AppId                = $App.appId
                    DisplayName          = $App.displayName
                    UserPrincipalName    = $null
                    Mail                 = $null
                    ServicePrincipalType = $null
                    AccountEnabled       = $null
                    ObjectType           = "ApplicationRegistration"
                    Status               = "Resolved"
                    ResolutionMessage    = $null
                }
            }
            else {

                throw (
                    "No service principal or app registration found " +
                    "for AppId $AppId"
                )
            }
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $null `
            -AppId $AppId `
            -ObjectType "Application" `
            -Message $_.Exception.Message
    }

    $ApplicationCache[$AppId] = $Result

    return $Result
}

# ============================================================
# Service Principal resolver
# ============================================================

function Resolve-CAServicePrincipal {

    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ($ServicePrincipalCache.ContainsKey($Id)) {

        $Stats.CacheHits++

        return $ServicePrincipalCache[$Id]
    }

    $Stats.ServicePrincipalLookups++

    try {

        $Uri =
            "https://graph.microsoft.com/v1.0/servicePrincipals/$Id" +
            '?$select=id,appId,displayName,servicePrincipalType,accountEnabled'

        $SP = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $Result = [PSCustomObject]@{
            Id                   = $SP.id
            AppId                = $SP.appId
            DisplayName          = $SP.displayName
            UserPrincipalName    = $null
            Mail                 = $null
            ServicePrincipalType = $SP.servicePrincipalType
            AccountEnabled       = $SP.accountEnabled
            ObjectType           = "ServicePrincipal"
            Status               = "Resolved"
            ResolutionMessage    = $null
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $Id `
            -AppId $null `
            -ObjectType "ServicePrincipal" `
            -Message $_.Exception.Message
    }

    $ServicePrincipalCache[$Id] = $Result

    return $Result
}

# ============================================================
# Role resolver
# ============================================================

function Resolve-CARole {

    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ($RoleCache.ContainsKey($Id)) {

        $Stats.CacheHits++

        return $RoleCache[$Id]
    }

    $Stats.RoleLookups++

    try {

        $Uri =
            "https://graph.microsoft.com/v1.0/directoryRoleTemplates/$Id"

        $Role = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $Result = [PSCustomObject]@{
            Id                = $Role.id
            AppId             = $null
            DisplayName       = $Role.displayName
            UserPrincipalName = $null
            Mail              = $null
            Description       = $Role.description
            ObjectType        = "DirectoryRole"
            Status            = "Resolved"
            ResolutionMessage = $null
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $Id `
            -AppId $null `
            -ObjectType "DirectoryRole" `
            -Message $_.Exception.Message
    }

    $RoleCache[$Id] = $Result

    return $Result
}

# ============================================================
# Location resolver
# ============================================================

function Resolve-CALocation {

    param (
        [Parameter(Mandatory)]
        [string]$Id
    )

    $Special = Get-SpecialCAValue -Value $Id

    if ($Special) {
        return $Special
    }

    if ($LocationCache.ContainsKey($Id)) {

        $Stats.CacheHits++

        return $LocationCache[$Id]
    }

    $Stats.LocationLookups++

    try {

        $Uri =
            "https://graph.microsoft.com/v1.0/identity/" +
            "conditionalAccess/namedLocations/$Id"

        $Location = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -ErrorAction Stop

        $LocationType = switch (
            $Location.'@odata.type'
        ) {

            "#microsoft.graph.ipNamedLocation" {
                "IP Named Location"
            }

            "#microsoft.graph.countryNamedLocation" {
                "Country Named Location"
            }

            default {
                "Named Location"
            }
        }

        $Result = [PSCustomObject]@{
            Id                = $Location.id
            AppId             = $null
            DisplayName       = $Location.displayName
            UserPrincipalName = $null
            Mail              = $null
            LocationType      = $LocationType
            IsTrusted         = $Location.isTrusted
            ObjectType        = "NamedLocation"
            Status            = "Resolved"
            ResolutionMessage = $null
        }
    }
    catch {

        $Stats.ResolutionFailures++

        $Result = New-UnresolvedResult `
            -Id $Id `
            -AppId $null `
            -ObjectType "NamedLocation" `
            -Message $_.Exception.Message
    }

    $LocationCache[$Id] = $Result

    return $Result
}

# ============================================================
# Detailed report row
# ============================================================

function Add-ReportRow {

    param (
        [string]$PolicyName,
        [string]$PolicyId,
        [string]$PolicyState,
        [string]$Category,
        [string]$Assignment,
        [string]$OriginalValue,
        [object]$ResolvedObject
    )

    $ReportRows.Add(
        [PSCustomObject]@{
            PolicyName        = $PolicyName
            PolicyId          = $PolicyId
            PolicyState       = $PolicyState
            Category          = $Category
            Assignment        = $Assignment
            ObjectType        = $ResolvedObject.ObjectType
            DisplayName       = $ResolvedObject.DisplayName
            OriginalValue     = $OriginalValue
            ObjectId          = $ResolvedObject.Id
            AppId             = $ResolvedObject.AppId
            UserPrincipalName = $ResolvedObject.UserPrincipalName
            Mail              = $ResolvedObject.Mail
            Status            = $ResolvedObject.Status
            ResolutionMessage = $ResolvedObject.ResolutionMessage
        }
    )
}

# ============================================================
# Collection resolver
# ============================================================

function Resolve-Collection {

    param (
        [object[]]$Values,

        [ValidateSet(
            "User",
            "Group",
            "Application",
            "ServicePrincipal",
            "Role",
            "Location"
        )]
        [string]$Type,

        [string]$PolicyName,
        [string]$PolicyId,
        [string]$PolicyState,
        [string]$Category,
        [string]$Assignment
    )

    $Results =
        [System.Collections.Generic.List[object]]::new()

    foreach ($Value in @($Values)) {

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$Value
            )
        ) {
            continue
        }

        $Resolved = switch ($Type) {

            "User" {
                Resolve-CAUser -Id $Value
            }

            "Group" {
                Resolve-CAGroup -Id $Value
            }

            "Application" {
                Resolve-CAApplication -AppId $Value
            }

            "ServicePrincipal" {
                Resolve-CAServicePrincipal -Id $Value
            }

            "Role" {
                Resolve-CARole -Id $Value
            }

            "Location" {
                Resolve-CALocation -Id $Value
            }
        }

        $Results.Add($Resolved)

        Add-ReportRow `
            -PolicyName $PolicyName `
            -PolicyId $PolicyId `
            -PolicyState $PolicyState `
            -Category $Category `
            -Assignment $Assignment `
            -OriginalValue ([string]$Value) `
            -ResolvedObject $Resolved
    }

    return @($Results)
}

function Convert-ResolvedCollectionToSummaryText {

    param (
        [object[]]$Objects
    )

    $Values = foreach ($Object in @($Objects)) {

        if ($null -eq $Object) {
            continue
        }

        if (
            -not [string]::IsNullOrWhiteSpace(
                [string]$Object.DisplayName
            )
        ) {

            $Object.DisplayName
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace(
                [string]$Object.UserPrincipalName
            )
        ) {

            $Object.UserPrincipalName
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace(
                [string]$Object.AppId
            )
        ) {

            $Object.AppId
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace(
                [string]$Object.Id
            )
        ) {

            $Object.Id
        }
    }

    return ($Values -join "; ")
}

function Get-UnresolvedCount {

    param (
        [object[]]$Collections
    )

    $Count = 0

    foreach ($Collection in @($Collections)) {

        foreach ($Object in @($Collection)) {

            if (
                $null -ne $Object -and
                $Object.Status -eq
                "PossiblyDeletedOrUnavailable"
            ) {

                $Count++
            }
        }
    }

    return $Count
}

# ============================================================
# Process freshly exported JSON files
# ============================================================

$JsonFiles = @(
    Get-ChildItem `
        -Path $OriginalJsonFolder `
        -Filter "*.json" `
        -File
)

Write-Host ""
Write-Host "Resolving Conditional Access Assignments" `
    -ForegroundColor Cyan
Write-Host "========================================" `
    -ForegroundColor Cyan
Write-Host ""

$PolicyNumber = 0

foreach ($File in $JsonFiles) {

    $PolicyNumber++

    Write-Progress `
        -Activity "Resolving Conditional Access Policies" `
        -Status (
            "Policy {0} of {1}: {2}" -f
            $PolicyNumber,
            $JsonFiles.Count,
            $File.BaseName
        ) `
        -PercentComplete (
            ($PolicyNumber / $JsonFiles.Count) * 100
        )

    try {

        $Policy = Get-Content `
            -Path $File.FullName `
            -Raw `
            -ErrorAction Stop |
        ConvertFrom-Json `
            -ErrorAction Stop
    }
    catch {

        $Stats.PoliciesFailed++

        Write-Warning (
            "Unable to parse policy file: " +
            $File.FullName
        )

        continue
    }

    $Stats.PoliciesProcessed++

    $PolicyName  = $Policy.DisplayName
    $PolicyId    = $Policy.Id
    $PolicyState = $Policy.State

    Write-Host (
        "[{0}/{1}] {2}" -f
        $PolicyNumber,
        $JsonFiles.Count,
        $PolicyName
    )

    $ResolvedPolicy =
        $Policy |
        ConvertTo-Json -Depth 100 |
        ConvertFrom-Json

    # --------------------------------------------------------
    # Per-policy collections
    # --------------------------------------------------------

    $IncludeUsers             = @()
    $ExcludeUsers             = @()
    $IncludeGroups            = @()
    $ExcludeGroups            = @()
    $IncludeRoles             = @()
    $ExcludeRoles             = @()
    $IncludeApplications      = @()
    $ExcludeApplications      = @()
    $IncludeServicePrincipals = @()
    $ExcludeServicePrincipals = @()
    $IncludeLocations         = @()
    $ExcludeLocations         = @()

    # --------------------------------------------------------
    # Users / Groups / Roles
    # --------------------------------------------------------

    if ($ResolvedPolicy.Conditions.Users) {

        $IncludeUsers = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.IncludeUsers `
                -Type User `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Users" `
                -Assignment "Include"
        )

        $ExcludeUsers = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.ExcludeUsers `
                -Type User `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Users" `
                -Assignment "Exclude"
        )

        $IncludeGroups = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.IncludeGroups `
                -Type Group `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Groups" `
                -Assignment "Include"
        )

        $ExcludeGroups = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.ExcludeGroups `
                -Type Group `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Groups" `
                -Assignment "Exclude"
        )

        $IncludeRoles = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.IncludeRoles `
                -Type Role `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Roles" `
                -Assignment "Include"
        )

        $ExcludeRoles = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Users.ExcludeRoles `
                -Type Role `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Roles" `
                -Assignment "Exclude"
        )

        $ResolvedPolicy.Conditions.Users.IncludeUsers =
            @($IncludeUsers)

        $ResolvedPolicy.Conditions.Users.ExcludeUsers =
            @($ExcludeUsers)

        $ResolvedPolicy.Conditions.Users.IncludeGroups =
            @($IncludeGroups)

        $ResolvedPolicy.Conditions.Users.ExcludeGroups =
            @($ExcludeGroups)

        $ResolvedPolicy.Conditions.Users.IncludeRoles =
            @($IncludeRoles)

        $ResolvedPolicy.Conditions.Users.ExcludeRoles =
            @($ExcludeRoles)
    }

    # --------------------------------------------------------
    # Applications
    # --------------------------------------------------------

    if ($ResolvedPolicy.Conditions.Applications) {

        $IncludeApplications = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Applications.IncludeApplications `
                -Type Application `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Applications" `
                -Assignment "Include"
        )

        $ExcludeApplications = @(
            Resolve-Collection `
                -Values $Policy.Conditions.Applications.ExcludeApplications `
                -Type Application `
                -PolicyName $PolicyName `
                -PolicyId $PolicyId `
                -PolicyState $PolicyState `
                -Category "Applications" `
                -Assignment "Exclude"
        )

        $ResolvedPolicy.Conditions.Applications.IncludeApplications =
            @($IncludeApplications)

        $ResolvedPolicy.Conditions.Applications.ExcludeApplications =
            @($ExcludeApplications)
    }

    # --------------------------------------------------------
    # Service Principals
    # --------------------------------------------------------

    if ($ResolvedPolicy.Conditions.ClientApplications) {

        if (
            $null -ne
            $Policy.Conditions.ClientApplications.IncludeServicePrincipals
        ) {

            $IncludeServicePrincipals = @(
                Resolve-Collection `
                    -Values $Policy.Conditions.ClientApplications.IncludeServicePrincipals `
                    -Type ServicePrincipal `
                    -PolicyName $PolicyName `
                    -PolicyId $PolicyId `
                    -PolicyState $PolicyState `
                    -Category "ServicePrincipals" `
                    -Assignment "Include"
            )

            $ResolvedPolicy.Conditions.ClientApplications.IncludeServicePrincipals =
                @($IncludeServicePrincipals)
        }

        if (
            $null -ne
            $Policy.Conditions.ClientApplications.ExcludeServicePrincipals
        ) {

            $ExcludeServicePrincipals = @(
                Resolve-Collection `
                    -Values $Policy.Conditions.ClientApplications.ExcludeServicePrincipals `
                    -Type ServicePrincipal `
                    -PolicyName $PolicyName `
                    -PolicyId $PolicyId `
                    -PolicyState $PolicyState `
                    -Category "ServicePrincipals" `
                    -Assignment "Exclude"
            )

            $ResolvedPolicy.Conditions.ClientApplications.ExcludeServicePrincipals =
                @($ExcludeServicePrincipals)
        }
    }

    # --------------------------------------------------------
    # Locations
    # --------------------------------------------------------

    if ($ResolvedPolicy.Conditions.Locations) {

        if (
            $null -ne
            $Policy.Conditions.Locations.IncludeLocations
        ) {

            $IncludeLocations = @(
                Resolve-Collection `
                    -Values $Policy.Conditions.Locations.IncludeLocations `
                    -Type Location `
                    -PolicyName $PolicyName `
                    -PolicyId $PolicyId `
                    -PolicyState $PolicyState `
                    -Category "Locations" `
                    -Assignment "Include"
            )

            $ResolvedPolicy.Conditions.Locations.IncludeLocations =
                @($IncludeLocations)
        }

        if (
            $null -ne
            $Policy.Conditions.Locations.ExcludeLocations
        ) {

            $ExcludeLocations = @(
                Resolve-Collection `
                    -Values $Policy.Conditions.Locations.ExcludeLocations `
                    -Type Location `
                    -PolicyName $PolicyName `
                    -PolicyId $PolicyId `
                    -PolicyState $PolicyState `
                    -Category "Locations" `
                    -Assignment "Exclude"
            )

            $ResolvedPolicy.Conditions.Locations.ExcludeLocations =
                @($ExcludeLocations)
        }
    }

    # --------------------------------------------------------
    # Other summary values
    # --------------------------------------------------------

    $GrantControls = @(
        $Policy.GrantControls.BuiltInControls
    ) -join "; "

    $AuthenticationStrength =
        $Policy.GrantControls.AuthenticationStrength.DisplayName

    $ClientAppTypes = @(
        $Policy.Conditions.ClientAppTypes
    ) -join "; "

    $PlatformsIncluded = @(
        $Policy.Conditions.Platforms.IncludePlatforms
    ) -join "; "

    $PlatformsExcluded = @(
        $Policy.Conditions.Platforms.ExcludePlatforms
    ) -join "; "

    $SignInRiskLevels = @(
        $Policy.Conditions.SignInRiskLevels
    ) -join "; "

    $UserRiskLevels = @(
        $Policy.Conditions.UserRiskLevels
    ) -join "; "

    $ServicePrincipalRiskLevels = @(
        $Policy.Conditions.ServicePrincipalRiskLevels
    ) -join "; "

    $IncludeUserActions = @(
        $Policy.Conditions.Applications.IncludeUserActions
    ) -join "; "

    $AuthenticationContexts = @(
        $Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences
    ) -join "; "

    $UnresolvedObjectCount = Get-UnresolvedCount `
        -Collections @(
            $IncludeUsers,
            $ExcludeUsers,
            $IncludeGroups,
            $ExcludeGroups,
            $IncludeRoles,
            $ExcludeRoles,
            $IncludeApplications,
            $ExcludeApplications,
            $IncludeServicePrincipals,
            $ExcludeServicePrincipals,
            $IncludeLocations,
            $ExcludeLocations
        )

    # --------------------------------------------------------
    # Policy summary
    # --------------------------------------------------------

    $PolicySummaryRows.Add(
        [PSCustomObject]@{
            PolicyName =
                $PolicyName

            PolicyId =
                $PolicyId

            State =
                $PolicyState

            CreatedDateTime =
                $Policy.CreatedDateTime

            ModifiedDateTime =
                $Policy.ModifiedDateTime

            Description =
                $Policy.Description

            IncludeUsers =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeUsers

            ExcludeUsers =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeUsers

            IncludeGroups =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeGroups

            ExcludeGroups =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeGroups

            IncludeRoles =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeRoles

            ExcludeRoles =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeRoles

            IncludeApplications =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeApplications

            ExcludeApplications =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeApplications

            IncludeServicePrincipals =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeServicePrincipals

            ExcludeServicePrincipals =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeServicePrincipals

            IncludeLocations =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $IncludeLocations

            ExcludeLocations =
                Convert-ResolvedCollectionToSummaryText `
                    -Objects $ExcludeLocations

            ClientAppTypes =
                $ClientAppTypes

            IncludePlatforms =
                $PlatformsIncluded

            ExcludePlatforms =
                $PlatformsExcluded

            SignInRiskLevels =
                $SignInRiskLevels

            UserRiskLevels =
                $UserRiskLevels

            ServicePrincipalRiskLevels =
                $ServicePrincipalRiskLevels

            IncludeUserActions =
                $IncludeUserActions

            AuthenticationContexts =
                $AuthenticationContexts

            GrantControls =
                $GrantControls

            GrantOperator =
                $Policy.GrantControls.Operator

            AuthenticationStrength =
                $AuthenticationStrength

            IncludedUserCount =
                $IncludeUsers.Count

            ExcludedUserCount =
                $ExcludeUsers.Count

            IncludedGroupCount =
                $IncludeGroups.Count

            ExcludedGroupCount =
                $ExcludeGroups.Count

            IncludedRoleCount =
                $IncludeRoles.Count

            ExcludedRoleCount =
                $ExcludeRoles.Count

            IncludedApplicationCount =
                $IncludeApplications.Count

            ExcludedApplicationCount =
                $ExcludeApplications.Count

            IncludedServicePrincipalCount =
                $IncludeServicePrincipals.Count

            ExcludedServicePrincipalCount =
                $ExcludeServicePrincipals.Count

            IncludedLocationCount =
                $IncludeLocations.Count

            ExcludedLocationCount =
                $ExcludeLocations.Count

            UnresolvedObjectCount =
                $UnresolvedObjectCount
        }
    )

    # --------------------------------------------------------
    # Resolved JSON
    # --------------------------------------------------------

    $ResolvedPath = Join-Path `
        $ResolvedJsonFolder `
        "$($File.BaseName)-Resolved.json"

    $ResolvedPolicy |
        ConvertTo-Json -Depth 100 |
        Set-Content `
            -Path $ResolvedPath `
            -Encoding UTF8

    if ($UnresolvedObjectCount -gt 0) {

        Write-Host (
            "    Unresolved objects: {0}" -f
            $UnresolvedObjectCount
        ) `
            -ForegroundColor Yellow
    }
}

Write-Progress `
    -Activity "Resolving Conditional Access Policies" `
    -Completed

# ============================================================
# Export reports
# ============================================================

if ($ReportRows.Count -gt 0) {

    $ReportRows |
        Sort-Object `
            PolicyName,
            Category,
            Assignment,
            DisplayName,
            OriginalValue |
        Export-Csv `
            -Path $DetailedCsvPath `
            -NoTypeInformation `
            -Encoding UTF8
}

if ($PolicySummaryRows.Count -gt 0) {

    $PolicySummaryRows |
        Sort-Object PolicyName |
        Export-Csv `
            -Path $PolicySummaryCsvPath `
            -NoTypeInformation `
            -Encoding UTF8
}

$TotalStopwatch.Stop()

# ============================================================
# Final summary
# ============================================================

Write-Host ""
Write-Host "Export and Resolution Complete" `
    -ForegroundColor Green
Write-Host "==============================" `
    -ForegroundColor Green
Write-Host ""

Write-Host "Policies exported            : $($Stats.PoliciesExported)"
Write-Host "Policies processed           : $($Stats.PoliciesProcessed)"
Write-Host "Policies failed              : $($Stats.PoliciesFailed)"
Write-Host "Detailed assignment rows     : $($ReportRows.Count)"
Write-Host "Policy summary rows           : $($PolicySummaryRows.Count)"
Write-Host ""

Write-Host "Graph Lookups" `
    -ForegroundColor Cyan
Write-Host "-------------"
Write-Host "Users                        : $($Stats.UserLookups)"
Write-Host "Groups                       : $($Stats.GroupLookups)"
Write-Host "Applications                 : $($Stats.ApplicationLookups)"
Write-Host "Service Principals           : $($Stats.ServicePrincipalLookups)"
Write-Host "Directory Roles              : $($Stats.RoleLookups)"
Write-Host "Named Locations              : $($Stats.LocationLookups)"
Write-Host "Cache hits                   : $($Stats.CacheHits)"
Write-Host "Resolution failures          : $($Stats.ResolutionFailures)"
Write-Host ""

Write-Host "Total elapsed                : $($TotalStopwatch.Elapsed)"
Write-Host ""

Write-Host "Original JSON exports:"
Write-Host "  $OriginalJsonFolder"
Write-Host ""

Write-Host "Resolved JSON exports:"
Write-Host "  $ResolvedJsonFolder"
Write-Host ""

Write-Host "Detailed assignment CSV:"
Write-Host "  $DetailedCsvPath"
Write-Host ""

Write-Host "Policy summary CSV:"
Write-Host "  $PolicySummaryCsvPath"
Write-Host ""
