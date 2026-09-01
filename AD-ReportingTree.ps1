#requires -Version 5.1

# ============================================================
# REPORTING TREE - PRODUCTION CANDIDATE
# ============================================================
#
# Optimized architecture:
#   - Interactive manager selection with exact/partial LDAP search
#   - Recursive LDAP_MATCHING_RULE_IN_CHAIN hierarchy discovery
#   - Raw LDAP enrichment in batches via DirectorySearcher
#   - Automatic AD module fallback for missing/failed enrichment results
#   - Disabled-manager collapsing and QC reporting
#   - Buffered console tree and CSV exports
#

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    throw "The ActiveDirectory PowerShell module could not be loaded. Error: $($_.Exception.Message)"
}

try {
    Add-Type -AssemblyName System.DirectoryServices -ErrorAction Stop
}
catch {
    throw "System.DirectoryServices could not be loaded. Error: $($_.Exception.Message)"
}

# Overall elapsed time (includes interactive manager selection time)
$TotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ============================================================
# CONFIGURATION
# ============================================================

# Starting manager.
#
# Examples:
#   $Manager = "jsmith"
#   $Manager = "123456"
#   $Manager = "CN=Jane Smith,OU=Users,DC=contoso,DC=com"
#
$Manager = ""


# ------------------------------------------------------------
# Domain Controller
# ------------------------------------------------------------
#
# OPTION 1:
# Specify a DC manually for consistent results:
#
# $DomainController = "dc01.contoso.com"
#
# OPTION 2:
# Leave blank and the script will automatically discover
# an AD Web Services-capable domain controller.
#
$DomainController = ""


# ------------------------------------------------------------
# Output Files
# ------------------------------------------------------------

$ReportingCsvPath = ".\ReportingTree.csv"
$DisabledCsvPath  = ".\ReportingTree_DisabledAccounts.csv"


# ------------------------------------------------------------
# LDAP Paging
# ------------------------------------------------------------
#
# Get-ADUser defaults to 256 objects per page.
# 500 is a reasonable explicit value for this report.
#
# This matters mainly if an individual manager has a very
# large number of direct reports.
#
$ResultPageSize = 500

# Maximum number of manager-search matches returned for interactive selection.
# This protects broad Contains/Last Name searches from returning thousands of
# users in large directories. Exact searches are normally unaffected.
$ManagerSearchResultLimit = 100


# ------------------------------------------------------------
# Recursive Hierarchy Discovery
# ------------------------------------------------------------
#
# Uses Active Directory's LDAP_MATCHING_RULE_IN_CHAIN matching rule
# to retrieve the entire descendant reporting branch in one server-side
# LDAP query. The returned users are then indexed by their Manager DN
# locally, so hierarchy construction does not require one LDAP request
# per user.
#
# If the recursive query fails, the script automatically falls back to
# the previous one-manager-at-a-time LDAP traversal.
#
$UseRecursiveLDAPDiscovery = $true

# Runtime state. Set to $true only after the recursive query succeeds.
$RecursiveDiscoveryLoaded = $false


# ------------------------------------------------------------
# Enrichment Batch Size
# ------------------------------------------------------------
#
# Full reporting properties are retrieved in LDAP batches after
# hierarchy discovery. 100 is a conservative starting point for
# reducing AD round trips without creating excessively large filters.
#
$EnrichmentBatchSize = 100

# Validate tunable numeric settings before any AD work begins.
if ($ResultPageSize -lt 1) {
    throw "ResultPageSize must be greater than zero."
}

if ($EnrichmentBatchSize -lt 1) {
    throw "EnrichmentBatchSize must be greater than zero."
}

if ($ManagerSearchResultLimit -lt 1) {
    throw "ManagerSearchResultLimit must be greater than zero."
}


# ============================================================
# PROPERTY SETS
# ============================================================

# ------------------------------------------------------------
# Phase 1: Lightweight Discovery Properties
# ------------------------------------------------------------
#
# These are the only extra properties requested while walking
# the reporting hierarchy.
#
# Keep this list small because these queries happen repeatedly.

$TraversalProperties = @(
    "DisplayName"
    "Manager"
    "Enabled"
)


# ------------------------------------------------------------
# Phase 2: Full Reporting Properties
# ------------------------------------------------------------
#
# These are retrieved ONLY for accounts actually discovered
# in the selected reporting hierarchy.

$ReportingProperties = @(
    "DisplayName"
    "UserPrincipalName"
    "Mail"
    "Title"
    "Department"
    "EmployeeID"
    "extensionAttribute1"
    "extensionAttribute2"
    "extensionAttribute3"
    "extensionAttribute4"
    "extensionAttribute5"
    "extensionAttribute6"
    "extensionAttribute7"
    "extensionAttribute8"
    "extensionAttribute9"
    "extensionAttribute10"
    "extensionAttribute11"
    "extensionAttribute12"
    "extensionAttribute13"
    "extensionAttribute14"
    "extensionAttribute15"
    "Enabled"
    "Manager"
)

# Raw LDAP attributes used by Step 9 enrichment.
# Enabled is derived from userAccountControl because LDAP does not expose
# the ActiveDirectory module's computed Enabled property directly.
$RawLdapReportingProperties = @(
    "distinguishedName"
    "displayName"
    "userPrincipalName"
    "sAMAccountName"
    "mail"
    "title"
    "department"
    "employeeID"
    "userAccountControl"
    "manager"
    "extensionAttribute1"
    "extensionAttribute2"
    "extensionAttribute3"
    "extensionAttribute4"
    "extensionAttribute5"
    "extensionAttribute6"
    "extensionAttribute7"
    "extensionAttribute8"
    "extensionAttribute9"
    "extensionAttribute10"
    "extensionAttribute11"
    "extensionAttribute12"
    "extensionAttribute13"
    "extensionAttribute14"
    "extensionAttribute15"
)


# ============================================================
# COLLECTIONS / CACHES
# ============================================================

# Main visible reporting output
$ReportingTree =
    [System.Collections.Generic.List[object]]::new()


# Disabled accounts encountered
$DisabledUsers =
    [System.Collections.Generic.List[object]]::new()

# Buffered console output. Building the tree in memory and writing it
# in one operation is substantially faster than one Write-Host call
# per user in larger reporting trees.
$ConsoleTreeLines = [System.Collections.Generic.List[string]]::new()


# ------------------------------------------------------------
# User Cache
# ------------------------------------------------------------
#
# DN -> lightweight AD user object
#
# Prevents unnecessary repeat lookups.

$UserCache =
    [System.Collections.Generic.Dictionary[string,object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


# ------------------------------------------------------------
# Direct Report Cache
# ------------------------------------------------------------
#
# Manager DN -> direct report collection
#
# This means the same manager DN is never queried twice.

$DirectReportCache =
    [System.Collections.Generic.Dictionary[string,object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


# ------------------------------------------------------------
# Enriched User Cache
# ------------------------------------------------------------
#
# DN -> full reporting user object
#
# Prevents duplicate enrichment calls.

$EnrichedUserCache =
    [System.Collections.Generic.Dictionary[string,object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


# ------------------------------------------------------------
# Report Data Cache
# ------------------------------------------------------------
#
# DN -> lightweight scalar report object. The normal production path
# is populated directly from raw LDAP SearchResult values. ADUser-based
# normalization is retained only as a fallback for resilience.

$NormalizedUserCache =
    [System.Collections.Generic.Dictionary[string,object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


# ------------------------------------------------------------
# Visited User Tracking
# ------------------------------------------------------------
#
# Protects against malformed / circular reporting data.

$VisitedUsers =
    [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )


# ============================================================
# SCRIPT STATISTICS
# ============================================================

$Stats = [ordered]@{

    LDAPQueries        = 0

    RecursiveDiscoveryQueries = 0

    RecursiveDiscoveryUsers   = 0

    RecursiveDiscoveryFallbacks = 0

    UsersDiscovered    = 0

    UsersEnriched      = 0

    UsersParsed        = 0

    EnrichmentBatches  = 0

    EnrichmentFallbacks = 0

    CacheHits          = 0

    CircularReferences = 0

    # Performance timing (milliseconds)
    DCDiscoveryMs      = 0.0
    ManagerSearchMs    = 0.0
    LDAPBranchMs       = 0.0
    EnrichmentMs       = 0.0
    EnrichmentLdapMs   = 0.0
    EnrichmentParseMs  = 0.0
    FallbackNormalizationMs = 0.0
    UsersNormalized    = 0
    NormalizationIdentityMs = 0.0
    NormalizationOrgMs      = 0.0
    NormalizationExtensionMs = 0.0
    NormalizationObjectMs   = 0.0
    HierarchyMs        = 0.0
    RenderReportMs     = 0.0
    VisibleTreeMs      = 0.0
    ReportObjectMs     = 0.0
    ReportCacheLookupMs = 0.0
    ReportDisplayPrepMs = 0.0
    ReportPSObjectMs    = 0.0
    ReportListAddMs     = 0.0
    ReportCacheFallbacks = 0
    ConsoleBuildMs     = 0.0
    ConsoleWriteMs     = 0.0
    DisabledReportMs   = 0.0
    CsvExportMs        = 0.0
}


# ============================================================
# LDAP FILTER ESCAPING
# ============================================================

function ConvertTo-LdapFilterValue {

    param (
        [Parameter(Mandatory)]
        [string]$Value
    )


    # RFC-style LDAP filter escaping.

    $Value = $Value -replace '\\', '\5c'
    $Value = $Value -replace '\*', '\2a'
    $Value = $Value -replace '\(', '\28'
    $Value = $Value -replace '\)', '\29'
    $Value = $Value -replace "`0", '\00'


    return $Value
}


# ============================================================
# INTERACTIVE MANAGER SELECTION
# ============================================================

function Select-StartingManager {

    Write-Host ""
    Write-Host "Starting Manager Selection"
    Write-Host "=========================="
    Write-Host ""
    Write-Host "Search using:"
    Write-Host "  1. UserPrincipalName (UPN)"
    Write-Host "  2. SamAccountName"
    Write-Host "  3. Mail"
    Write-Host "  4. DisplayName"
    Write-Host "  5. First Name"
    Write-Host "  6. Last Name"
    Write-Host ""

    do {
        $SearchType = Read-Host "Select search type [1-6]"
    }
    until ($SearchType -match '^[1-6]$')


    Write-Host ""
    Write-Host "Match type:"
    Write-Host "  1. Exact"
    Write-Host "  2. Starts with"
    Write-Host "  3. Contains"
    Write-Host ""

    do {
        $MatchType = Read-Host "Select match type [1-3]"
    }
    until ($MatchType -match '^[1-3]$')


    $SearchValue = Read-Host "Enter manager search value"

    if ([string]::IsNullOrWhiteSpace($SearchValue)) {
        throw "Manager search value cannot be blank."
    }


    # Escape the user's literal input before adding any wildcard
    # required by the selected match type. This prevents LDAP
    # special characters entered by the user from changing the
    # intended filter.
    $EscapedValue = ConvertTo-LdapFilterValue `
        -Value $SearchValue.Trim()


    switch ($SearchType) {

        "1" {
            $SearchLabel = "UserPrincipalName"
            $LDAPAttribute = "userPrincipalName"
        }

        "2" {
            $SearchLabel = "SamAccountName"
            $LDAPAttribute = "sAMAccountName"
        }

        "3" {
            $SearchLabel = "Mail"
            $LDAPAttribute = "mail"
        }

        "4" {
            $SearchLabel = "DisplayName"
            $LDAPAttribute = "displayName"
        }

        "5" {
            $SearchLabel = "First Name"
            $LDAPAttribute = "givenName"
        }

        "6" {
            $SearchLabel = "Last Name"
            $LDAPAttribute = "sn"
        }
    }


    switch ($MatchType) {

        "1" {
            $MatchLabel = "Exact"
            $LDAPSearchValue = $EscapedValue
        }

        "2" {
            $MatchLabel = "Starts with"
            $LDAPSearchValue = "$EscapedValue*"
        }

        "3" {
            $MatchLabel = "Contains"
            $LDAPSearchValue = "*$EscapedValue*"
        }
    }


    $LDAPFilter = "($LDAPAttribute=$LDAPSearchValue)"


    Write-Host ""
    Write-Host (
        "Searching $SearchLabel using '$MatchLabel' for '$SearchValue'..."
    )


    try {

        $ManagerSearchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $Matches = @(
                Get-ADUser `
                    -LDAPFilter $LDAPFilter `
                    -Server $DomainController `
                    -Properties `
                        DisplayName,
                        UserPrincipalName,
                        Mail,
                        GivenName,
                        Surname,
                        Enabled,
                        Manager `
                    -ResultPageSize $ResultPageSize `
                    -ResultSetSize $ManagerSearchResultLimit `
                    -ErrorAction Stop |
                Sort-Object DisplayName, SamAccountName
            )
        }
        finally {
            $ManagerSearchStopwatch.Stop()
            $script:Stats.ManagerSearchMs += $ManagerSearchStopwatch.Elapsed.TotalMilliseconds
        }
    }
    catch {

        throw (
            "Unable to search Active Directory. Error: " +
            $_.Exception.Message
        )
    }


    # --------------------------------------------------------
    # No Matches
    # --------------------------------------------------------

    if ($Matches.Count -eq 0) {

        Write-Warning (
            "No users were found where $SearchLabel " +
            "matches '$SearchValue' using '$MatchLabel'."
        )

        return $null
    }


    # --------------------------------------------------------
    # Exactly One Match
    # --------------------------------------------------------

    if ($Matches.Count -eq 1) {

        $SelectedUser = $Matches[0]

        Write-Host ""
        Write-Host "Manager found:"
        Write-Host ""
        Write-Host "  Display Name : $($SelectedUser.DisplayName)"
        Write-Host "  UPN          : $($SelectedUser.UserPrincipalName)"
        Write-Host "  SAM Account  : $($SelectedUser.SamAccountName)"
        Write-Host "  Mail         : $($SelectedUser.Mail)"
        Write-Host "  Enabled      : $($SelectedUser.Enabled)"
        Write-Host ""

        return $SelectedUser
    }


    # --------------------------------------------------------
    # Multiple Matches
    # --------------------------------------------------------

    if ($Matches.Count -ge $ManagerSearchResultLimit) {
        Write-Warning (
            "The search returned at least $ManagerSearchResultLimit users. " +
            "Refine the search value or use Exact/Starts with for a narrower result set."
        )
    }

    Write-Host ""
    Write-Host (
        "$($Matches.Count) users matched '$SearchValue' using '$MatchLabel'."
    )
    Write-Host ""
    Write-Host "Select the correct manager:"
    Write-Host ""


    for ($i = 0; $i -lt $Matches.Count; $i++) {

        $Match = $Matches[$i]

        Write-Host (
            "[$($i + 1)] " +
            "$($Match.DisplayName) | " +
            "$($Match.UserPrincipalName) | " +
            "$($Match.SamAccountName) | " +
            "$($Match.Mail) | " +
            "Enabled: $($Match.Enabled)"
        )
    }


    Write-Host ""
    Write-Host "[0] Cancel"
    Write-Host ""


    do {

        $Selection = Read-Host (
            "Select manager [0-$($Matches.Count)]"
        )

        $ValidSelection =
            $Selection -match '^\d+$' -and
            [int]$Selection -ge 0 -and
            [int]$Selection -le $Matches.Count

    }
    until ($ValidSelection)


    if ([int]$Selection -eq 0) {

        return $null
    }


    $SelectedUser =
        $Matches[[int]$Selection - 1]


    Write-Host ""
    Write-Host "Selected manager:"
    Write-Host ""
    Write-Host "  Display Name : $($SelectedUser.DisplayName)"
    Write-Host "  UPN          : $($SelectedUser.UserPrincipalName)"
    Write-Host "  SAM Account  : $($SelectedUser.SamAccountName)"
    Write-Host "  Mail         : $($SelectedUser.Mail)"
    Write-Host "  Enabled      : $($SelectedUser.Enabled)"
    Write-Host ""


    return $SelectedUser
}


# ============================================================
# SAFE DISPLAY NAME
# ============================================================

function Get-UserDisplayName {

    param (
        [Parameter(Mandatory)]
        [object]$User
    )


    if (
        -not [string]::IsNullOrWhiteSpace(
            $User.DisplayName
        )
    ) {

        return $User.DisplayName
    }


    if (
        -not [string]::IsNullOrWhiteSpace(
            $User.SamAccountName
        )
    ) {

        return $User.SamAccountName
    }


    return $User.DistinguishedName
}


# ============================================================
# DISCOVER DOMAIN CONTROLLER
# ============================================================

if ([string]::IsNullOrWhiteSpace($DomainController)) {

    Write-Host ""
    Write-Host "Discovering domain controller..."

    try {

        $DCStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $DiscoveredDC = Get-ADDomainController `
                -Discover `
                -Service ADWS `
                -ErrorAction Stop
        }
        finally {
            $DCStopwatch.Stop()
            $Stats.DCDiscoveryMs += $DCStopwatch.Elapsed.TotalMilliseconds
        }

        # HostName can be returned as an ADPropertyValueCollection.
        # Explicitly select the first value and cast it to string
        # for use with the -Server parameter.
        $DomainController = [string]($DiscoveredDC.HostName | Select-Object -First 1)

        if ([string]::IsNullOrWhiteSpace($DomainController)) {
            throw "Domain controller discovery returned an empty HostName."
        }

        Write-Host "Using domain controller: $DomainController"
    }
    catch {

        Write-Error (
            "Unable to discover an AD Web Services " +
            "domain controller. Error: " +
            $_.Exception.Message
        )

        return
    }
}
else {

    # Ensure a manually configured value is also a string.
    $DomainController = [string]$DomainController

    Write-Host ""
    Write-Host (
        "Using configured domain controller: " +
        $DomainController
    )
}


# ============================================================
# GET STARTING MANAGER
# ============================================================

$RootUser = $null


while ($null -eq $RootUser) {

    try {

        $RootUser = Select-StartingManager
    }
    catch {

        Write-Warning $_.Exception.Message
        $RootUser = $null
    }


    if ($null -eq $RootUser) {

        Write-Host ""
        $Retry = Read-Host "Search again? [Y/N]"

        if ($Retry -notmatch '^[Yy]$') {

            Write-Host ""
            Write-Host "Report cancelled."
            return
        }
    }
}


# Add selected manager to cache.

$UserCache[$RootUser.DistinguishedName] =
    $RootUser


# ============================================================
# RECURSIVE REPORTING-BRANCH DISCOVERY
# ============================================================
#
# LDAP_MATCHING_RULE_IN_CHAIN OID:
#   1.2.840.113556.1.4.1941
#
# For the DN-valued Manager attribute, this asks AD to return users whose
# manager chain eventually reaches the selected root manager. This allows
# the full descendant branch to be discovered with one LDAP query rather
# than one query per user.
#
# After retrieval, DirectReportCache is populated from each user's actual
# Manager attribute. Get-ActualReportingHierarchy can then build the same
# tree entirely from the local cache.

function Initialize-RecursiveReportingDiscovery {

    param (
        [Parameter(Mandatory)]
        [object]$RootUser
    )


    if (-not $UseRecursiveLDAPDiscovery) {
        return $false
    }


    $EscapedRootDN =
        ConvertTo-LdapFilterValue `
            -Value $RootUser.DistinguishedName


    $LDAPFilter =
        "(manager:1.2.840.113556.1.4.1941:=$EscapedRootDN)"


    Write-Host ""
    Write-Host "Preloading reporting branch with recursive LDAP discovery..."


    $RecursiveStopwatch =
        [System.Diagnostics.Stopwatch]::StartNew()


    try {

        $script:Stats.LDAPQueries++
        $script:Stats.RecursiveDiscoveryQueries++


        $Descendants = @(
            Get-ADUser `
                -LDAPFilter $LDAPFilter `
                -Server $DomainController `
                -Properties $TraversalProperties `
                -ResultPageSize $ResultPageSize `
                -ResultSetSize $null `
                -ErrorAction Stop
        )


        $RecursiveStopwatch.Stop()

        $script:Stats.LDAPBranchMs +=
            $RecursiveStopwatch.Elapsed.TotalMilliseconds

        $script:Stats.RecursiveDiscoveryUsers =
            $Descendants.Count


        # ----------------------------------------------------
        # Cache All Discovered Users
        # ----------------------------------------------------

        foreach ($User in $Descendants) {

            $UserCache[$User.DistinguishedName] =
                $User
        }


        # Root was retrieved during manager selection.
        $UserCache[$RootUser.DistinguishedName] =
            $RootUser


        # ----------------------------------------------------
        # Build Manager -> Direct Reports Index
        # ----------------------------------------------------
        #
        # Initialize a cache entry for every discovered user. This is
        # important for leaf users: an empty cached list tells
        # Get-DirectReports that no LDAP query is required.

        foreach ($User in $Descendants) {

            if (-not $DirectReportCache.ContainsKey(
                    $User.DistinguishedName
                )) {

                $DirectReportCache[
                    $User.DistinguishedName
                ] = @()
            }
        }


        if (-not $DirectReportCache.ContainsKey(
                $RootUser.DistinguishedName
            )) {

            $DirectReportCache[
                $RootUser.DistinguishedName
            ] = @()
        }


        # Temporary manager buckets make it inexpensive to append users.
        $ManagerBuckets =
            [System.Collections.Generic.Dictionary[string,object]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )


        foreach ($User in $Descendants) {

            if ([string]::IsNullOrWhiteSpace($User.Manager)) {
                continue
            }


            if (-not $ManagerBuckets.ContainsKey($User.Manager)) {

                $ManagerBuckets[$User.Manager] =
                    [System.Collections.Generic.List[object]]::new()
            }


            $ManagerBuckets[$User.Manager].Add($User)
        }


        foreach ($ManagerDN in $ManagerBuckets.Keys) {

            $DirectReportCache[$ManagerDN] = @(
                $ManagerBuckets[$ManagerDN] |
                    Sort-Object `
                        @{ Expression = { $_.DisplayName } },
                        @{ Expression = { $_.SamAccountName } }
            )
        }


        $script:RecursiveDiscoveryLoaded = $true


        Write-Host (
            "Recursive LDAP discovery loaded " +
            "$($Descendants.Count) descendant users in " +
            ("{0:N2} sec." -f (
                $RecursiveStopwatch.Elapsed.TotalSeconds
            ))
        )


        return $true
    }
    catch {

        if ($RecursiveStopwatch.IsRunning) {
            $RecursiveStopwatch.Stop()
        }


        $script:Stats.LDAPBranchMs +=
            $RecursiveStopwatch.Elapsed.TotalMilliseconds

        $script:Stats.RecursiveDiscoveryFallbacks++
        $script:RecursiveDiscoveryLoaded = $false


        # Clear any partial cache entries created by a failed attempt.
        $DirectReportCache.Clear()


        Write-Warning (
            "Recursive LDAP hierarchy discovery failed. " +
            "Falling back to per-manager LDAP traversal. Error: " +
            $_.Exception.Message
        )


        return $false
    }
}


# ============================================================
# DIRECT REPORT LDAP QUERY
# ============================================================

function Get-DirectReports {

    param (
        [Parameter(Mandatory)]
        [string]$ManagerDN
    )


    # --------------------------------------------------------
    # Check Cache
    # --------------------------------------------------------

    if (
        $DirectReportCache.ContainsKey(
            $ManagerDN
        )
    ) {

        $script:Stats.CacheHits++

        return @(
            $DirectReportCache[$ManagerDN]
        )
    }


    # --------------------------------------------------------
    # Recursive Discovery Leaf / Out-of-Branch Handling
    # --------------------------------------------------------
    #
    # When recursive discovery succeeded, all users in the selected
    # branch were preloaded and indexed. A missing manager key therefore
    # means there are no in-branch direct reports and no LDAP request is
    # needed.

    if ($script:RecursiveDiscoveryLoaded) {
        return @()
    }


    # --------------------------------------------------------
    # Escape Manager DN
    # --------------------------------------------------------

    $EscapedManagerDN =
        ConvertTo-LdapFilterValue `
            -Value $ManagerDN


    # --------------------------------------------------------
    # LDAP Query
    # --------------------------------------------------------
    #
    # Filtering occurs on the domain controller.
    #
    # Only users whose manager attribute exactly matches the
    # current manager DN are returned.

    $LDAPFilter =
        "(manager=$EscapedManagerDN)"


    try {

        $script:Stats.LDAPQueries++
        $LDAPStopwatch = [System.Diagnostics.Stopwatch]::StartNew()


        $DirectReports = @(
            Get-ADUser `
                -LDAPFilter $LDAPFilter `
                -Server $DomainController `
                -Properties $TraversalProperties `
                -ResultPageSize $ResultPageSize `
                -ResultSetSize $null `
                -ErrorAction Stop |
            Sort-Object `
                @{ Expression = {
                    $_.DisplayName
                } },
                @{ Expression = {
                    $_.SamAccountName
                } }
        )

        $LDAPStopwatch.Stop()
        $script:Stats.LDAPBranchMs += $LDAPStopwatch.Elapsed.TotalMilliseconds
    }
    catch {

        if ($null -ne $LDAPStopwatch -and $LDAPStopwatch.IsRunning) {
            $LDAPStopwatch.Stop()
            $script:Stats.LDAPBranchMs += $LDAPStopwatch.Elapsed.TotalMilliseconds
        }

        Write-Warning (
            "Unable to retrieve direct reports for DN: " +
            $ManagerDN +
            ". Error: " +
            $_.Exception.Message
        )


        $DirectReports = @()
    }


    # --------------------------------------------------------
    # Populate User Cache
    # --------------------------------------------------------

    foreach ($User in $DirectReports) {

        $UserCache[$User.DistinguishedName] =
            $User
    }


    # --------------------------------------------------------
    # Cache Direct Reports
    # --------------------------------------------------------

    $DirectReportCache[$ManagerDN] =
        $DirectReports


    return $DirectReports
}


# ============================================================
# ENRICH USER
# ============================================================

function Get-EnrichedUser {

    param (
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )


    # --------------------------------------------------------
    # Check Enrichment Cache
    # --------------------------------------------------------

    if (
        $EnrichedUserCache.ContainsKey(
            $DistinguishedName
        )
    ) {

        $script:Stats.CacheHits++

        return $EnrichedUserCache[
            $DistinguishedName
        ]
    }


    # --------------------------------------------------------
    # Retrieve Full Reporting Properties
    # --------------------------------------------------------

    try {

        $EnrichmentStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $User =
                Get-ADUser `
                    -Identity $DistinguishedName `
                    -Server $DomainController `
                    -Properties $ReportingProperties `
                    -ErrorAction Stop
        }
        finally {
            $EnrichmentStopwatch.Stop()
            $script:Stats.EnrichmentMs += $EnrichmentStopwatch.Elapsed.TotalMilliseconds
        }


        $script:Stats.UsersEnriched++


        $EnrichedUserCache[
            $DistinguishedName
        ] = $User


        return $User
    }
    catch {

        Write-Warning (
            "Unable to retrieve reporting properties for " +
            "'$DistinguishedName'. Error: " +
            $_.Exception.Message
        )


        # Fall back to lightweight cached object if possible.

        if (
            $UserCache.ContainsKey(
                $DistinguishedName
            )
        ) {

            return $UserCache[
                $DistinguishedName
            ]
        }


        return $null
    }
}


# ============================================================
# NORMALIZE ENRICHED AD USER - STEP 9 FALLBACK ADUSER NORMALIZATION
# ============================================================
#
# Step 9 uses raw LDAP for the production enrichment path. This function
# remains only for the rare Get-ADUser fallback path.
#
# IMPORTANT:
#   $User.PSObject.Properties itself is a normal PowerShell metadata
#   lookup. Individual AD values are then read through the property
#   collection instead of expressions such as $User.DisplayName.

function ConvertTo-NormalizedUser {

    param (
        [Parameter(Mandatory)]
        [object]$User
    )

    $Properties = $User.PSObject.Properties

    # --------------------------------------------------------
    # Identity / addressing properties
    # --------------------------------------------------------

    $GroupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $DistinguishedName = [string]$Properties['DistinguishedName'].Value
    $DisplayName       = [string]$Properties['DisplayName'].Value
    $UserPrincipalName = [string]$Properties['UserPrincipalName'].Value
    $SamAccountName    = [string]$Properties['SamAccountName'].Value
    $Mail              = [string]$Properties['Mail'].Value

    $GroupStopwatch.Stop()
    $script:Stats.NormalizationIdentityMs += $GroupStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Organizational / account properties
    # --------------------------------------------------------

    $GroupStopwatch.Restart()

    $Title      = [string]$Properties['Title'].Value
    $Department = [string]$Properties['Department'].Value
    $EmployeeID = [string]$Properties['EmployeeID'].Value
    $Enabled    = [bool]$Properties['Enabled'].Value
    $Manager    = [string]$Properties['Manager'].Value

    $GroupStopwatch.Stop()
    $script:Stats.NormalizationOrgMs += $GroupStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # extensionAttribute1-15
    # --------------------------------------------------------

    $GroupStopwatch.Restart()

    $extensionAttribute1  = [string]$Properties['extensionAttribute1'].Value
    $extensionAttribute2  = [string]$Properties['extensionAttribute2'].Value
    $extensionAttribute3  = [string]$Properties['extensionAttribute3'].Value
    $extensionAttribute4  = [string]$Properties['extensionAttribute4'].Value
    $extensionAttribute5  = [string]$Properties['extensionAttribute5'].Value
    $extensionAttribute6  = [string]$Properties['extensionAttribute6'].Value
    $extensionAttribute7  = [string]$Properties['extensionAttribute7'].Value
    $extensionAttribute8  = [string]$Properties['extensionAttribute8'].Value
    $extensionAttribute9  = [string]$Properties['extensionAttribute9'].Value
    $extensionAttribute10 = [string]$Properties['extensionAttribute10'].Value
    $extensionAttribute11 = [string]$Properties['extensionAttribute11'].Value
    $extensionAttribute12 = [string]$Properties['extensionAttribute12'].Value
    $extensionAttribute13 = [string]$Properties['extensionAttribute13'].Value
    $extensionAttribute14 = [string]$Properties['extensionAttribute14'].Value
    $extensionAttribute15 = [string]$Properties['extensionAttribute15'].Value

    $GroupStopwatch.Stop()
    $script:Stats.NormalizationExtensionMs += $GroupStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Create lightweight scalar-only object
    # --------------------------------------------------------

    $GroupStopwatch.Restart()

    $Normalized = [PSCustomObject][ordered]@{
        DistinguishedName    = $DistinguishedName
        DisplayName          = $DisplayName
        UserPrincipalName    = $UserPrincipalName
        SamAccountName       = $SamAccountName
        Mail                 = $Mail
        Title                = $Title
        Department           = $Department
        EmployeeID           = $EmployeeID
        Enabled              = $Enabled
        Manager              = $Manager
        extensionAttribute1  = $extensionAttribute1
        extensionAttribute2  = $extensionAttribute2
        extensionAttribute3  = $extensionAttribute3
        extensionAttribute4  = $extensionAttribute4
        extensionAttribute5  = $extensionAttribute5
        extensionAttribute6  = $extensionAttribute6
        extensionAttribute7  = $extensionAttribute7
        extensionAttribute8  = $extensionAttribute8
        extensionAttribute9  = $extensionAttribute9
        extensionAttribute10 = $extensionAttribute10
        extensionAttribute11 = $extensionAttribute11
        extensionAttribute12 = $extensionAttribute12
        extensionAttribute13 = $extensionAttribute13
        extensionAttribute14 = $extensionAttribute14
        extensionAttribute15 = $extensionAttribute15
    }

    $GroupStopwatch.Stop()
    $script:Stats.NormalizationObjectMs += $GroupStopwatch.Elapsed.TotalMilliseconds

    return $Normalized
}


# ============================================================
# RAW LDAP ENRICHMENT HELPERS
# ============================================================

function Get-RawLdapPropertyValue {
    param (
        [Parameter(Mandatory)]
        [System.DirectoryServices.SearchResult]$Result,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $Values = $Result.Properties[$Name]

    if ($null -eq $Values -or $Values.Count -eq 0) {
        return $null
    }

    return $Values[0]
}


function ConvertFrom-RawLdapResult {
    param (
        [Parameter(Mandatory)]
        [System.DirectoryServices.SearchResult]$Result
    )

    $ParseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $DistinguishedName = [string](Get-RawLdapPropertyValue -Result $Result -Name 'distinguishedName')
        $DisplayName = [string](Get-RawLdapPropertyValue -Result $Result -Name 'displayName')
        $UserPrincipalName = [string](Get-RawLdapPropertyValue -Result $Result -Name 'userPrincipalName')
        $SamAccountName = [string](Get-RawLdapPropertyValue -Result $Result -Name 'sAMAccountName')
        $Mail = [string](Get-RawLdapPropertyValue -Result $Result -Name 'mail')
        $Title = [string](Get-RawLdapPropertyValue -Result $Result -Name 'title')
        $Department = [string](Get-RawLdapPropertyValue -Result $Result -Name 'department')
        $EmployeeID = [string](Get-RawLdapPropertyValue -Result $Result -Name 'employeeID')
        $Manager = [string](Get-RawLdapPropertyValue -Result $Result -Name 'manager')

        $UserAccountControlValue = Get-RawLdapPropertyValue -Result $Result -Name 'userAccountControl'
        $UserAccountControl = if ($null -eq $UserAccountControlValue) { 0 } else { [int64]$UserAccountControlValue }
        $Enabled = (($UserAccountControl -band 2) -eq 0)

        $Normalized = [PSCustomObject][ordered]@{
            DistinguishedName    = $DistinguishedName
            DisplayName          = $DisplayName
            UserPrincipalName    = $UserPrincipalName
            SamAccountName       = $SamAccountName
            Mail                 = $Mail
            Title                = $Title
            Department           = $Department
            EmployeeID           = $EmployeeID
            Enabled              = $Enabled
            Manager              = $Manager
            extensionAttribute1  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute1')
            extensionAttribute2  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute2')
            extensionAttribute3  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute3')
            extensionAttribute4  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute4')
            extensionAttribute5  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute5')
            extensionAttribute6  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute6')
            extensionAttribute7  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute7')
            extensionAttribute8  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute8')
            extensionAttribute9  = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute9')
            extensionAttribute10 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute10')
            extensionAttribute11 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute11')
            extensionAttribute12 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute12')
            extensionAttribute13 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute13')
            extensionAttribute14 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute14')
            extensionAttribute15 = [string](Get-RawLdapPropertyValue -Result $Result -Name 'extensionAttribute15')
        }

        return $Normalized
    }
    finally {
        $ParseStopwatch.Stop()
        $script:Stats.EnrichmentParseMs += $ParseStopwatch.Elapsed.TotalMilliseconds
    }
}


function New-RawLdapSearcher {
    param (
        [Parameter(Mandatory)]
        [string]$LDAPFilter
    )

    # Bind directly to the selected DC's default naming context so the raw
    # enrichment query observes the same directory server as the AD cmdlets.
    $RootDse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController/RootDSE")
    $DefaultNamingContext = [string]$RootDse.Properties['defaultNamingContext'][0]
    $RootDse.Dispose()

    if ([string]::IsNullOrWhiteSpace($DefaultNamingContext)) {
        throw "Unable to determine defaultNamingContext from $DomainController."
    }

    $SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController/$DefaultNamingContext")
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher($SearchRoot)

    $Searcher.Filter = $LDAPFilter
    $Searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
    $Searcher.PageSize = $ResultPageSize
    $Searcher.SizeLimit = 0
    $Searcher.CacheResults = $false

    $Searcher.PropertiesToLoad.Clear()
    foreach ($PropertyName in $RawLdapReportingProperties) {
        [void]$Searcher.PropertiesToLoad.Add($PropertyName)
    }

    return $Searcher
}


# ============================================================
# BATCH ENRICHMENT
# ============================================================
#
# After hierarchy discovery, retrieve full reporting properties for
# discovered users in batches. This replaces one enrichment AD request
# per user with roughly one request per $EnrichmentBatchSize users.
#
# Get-EnrichedUser remains available as a fallback if an individual
# batch fails or a returned batch is missing a discovered account.

function Initialize-BatchedEnrichment {

    param (
        [Parameter(Mandatory)]
        [object]$RootNode
    )

    $DiscoveredDNs = [System.Collections.Generic.List[string]]::new()

    function Add-NodeDNs {
        param (
            [Parameter(Mandatory)]
            [object]$Node
        )

        $DiscoveredDNs.Add([string]$Node.DistinguishedName)

        foreach ($Child in $Node.Children) {
            Add-NodeDNs -Node $Child
        }
    }

    Add-NodeDNs -Node $RootNode

    if ($DiscoveredDNs.Count -eq 0) {
        return
    }

    Write-Host ""
    Write-Host (
        "Batch enriching $($DiscoveredDNs.Count) discovered users " +
        "with raw LDAP (batch size: $EnrichmentBatchSize)..."
    )

    $EnrichmentTotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        for (
            $Start = 0;
            $Start -lt $DiscoveredDNs.Count;
            $Start += $EnrichmentBatchSize
        ) {
            $End = [Math]::Min(
                $Start + $EnrichmentBatchSize - 1,
                $DiscoveredDNs.Count - 1
            )

            $BatchDNs = @($DiscoveredDNs[$Start..$End])

            $FilterParts = foreach ($DN in $BatchDNs) {
                $EscapedDN = ConvertTo-LdapFilterValue -Value $DN
                "(distinguishedName=$EscapedDN)"
            }

            $LDAPFilter = "(|" + ($FilterParts -join "") + ")"
            $script:Stats.EnrichmentBatches++

            $ReturnedDNs = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )

            $Searcher = $null
            $Results = $null
            $BatchSucceeded = $false

            try {
                $Searcher = New-RawLdapSearcher -LDAPFilter $LDAPFilter

                $LdapStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    $Results = $Searcher.FindAll()
                }
                finally {
                    $LdapStopwatch.Stop()
                    $script:Stats.EnrichmentLdapMs += $LdapStopwatch.Elapsed.TotalMilliseconds
                }

                foreach ($Result in $Results) {
                    $NormalizedUser = ConvertFrom-RawLdapResult -Result $Result

                    if ($null -eq $NormalizedUser -or [string]::IsNullOrWhiteSpace($NormalizedUser.DistinguishedName)) {
                        continue
                    }

                    $NormalizedUserCache[$NormalizedUser.DistinguishedName] = $NormalizedUser
                    $EnrichedUserCache[$NormalizedUser.DistinguishedName] = $NormalizedUser
                    [void]$ReturnedDNs.Add($NormalizedUser.DistinguishedName)

                    $script:Stats.UsersEnriched++
                    $script:Stats.UsersParsed++
                }

                $BatchSucceeded = $true
            }
            catch {
                Write-Warning (
                    "Raw LDAP enrichment batch $($script:Stats.EnrichmentBatches) failed. " +
                    "Falling back to individual Get-ADUser enrichment for " +
                    "$($BatchDNs.Count) users. Error: $($_.Exception.Message)"
                )
            }
            finally {
                if ($null -ne $Results) {
                    $Results.Dispose()
                }

                if ($null -ne $Searcher) {
                    $SearchRootToDispose = $Searcher.SearchRoot
                    $Searcher.Dispose()
                    if ($null -ne $SearchRootToDispose) {
                        $SearchRootToDispose.Dispose()
                    }
                }
            }

            # If raw LDAP failed, or if a requested DN was not returned, retain
            # the existing individual AD-module fallback for reliability.
            foreach ($DN in $BatchDNs) {
                if (-not $BatchSucceeded -or -not $ReturnedDNs.Contains($DN)) {
                    $script:Stats.EnrichmentFallbacks++

                    $FallbackUser = Get-EnrichedUser -DistinguishedName $DN

                    if ($null -ne $FallbackUser -and -not $NormalizedUserCache.ContainsKey($DN)) {
                        $NormalizationStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                        $NormalizedUser = ConvertTo-NormalizedUser -User $FallbackUser
                        $NormalizationStopwatch.Stop()

                        $script:Stats.FallbackNormalizationMs += $NormalizationStopwatch.Elapsed.TotalMilliseconds
                        $script:Stats.UsersNormalized++

                        if ($null -ne $NormalizedUser) {
                            $NormalizedUserCache[$DN] = $NormalizedUser
                        }
                    }
                }
            }
        }
    }
    finally {
        $EnrichmentTotalStopwatch.Stop()
        $script:Stats.EnrichmentMs += $EnrichmentTotalStopwatch.Elapsed.TotalMilliseconds
    }

    Write-Host (
        "Raw LDAP enrichment complete: $($NormalizedUserCache.Count) users cached."
    )

    if ($NormalizedUserCache.Count -lt $DiscoveredDNs.Count) {
        Write-Warning (
            "Enrichment completed with fewer cached users than discovered. " +
            "Discovered: $($DiscoveredDNs.Count); Cached: $($NormalizedUserCache.Count). " +
            "Review fallback warnings and the exported report."
        )
    }
}


# ============================================================
# RESOLVE MANAGER DISPLAY NAME
# ============================================================

function Get-ManagerDisplayName {

    param (
        [string]$ManagerDN
    )


    if (
        [string]::IsNullOrWhiteSpace(
            $ManagerDN
        )
    ) {

        return ""
    }


    # --------------------------------------------------------
    # Try Local Cache First
    # --------------------------------------------------------

    if (
        $UserCache.ContainsKey(
            $ManagerDN
        )
    ) {

        $script:Stats.CacheHits++


        return Get-UserDisplayName `
            -User $UserCache[$ManagerDN]
    }


    # --------------------------------------------------------
    # Retrieve Manager Only If Necessary
    # --------------------------------------------------------

    try {

        $ManagerObject =
            Get-ADUser `
                -Identity $ManagerDN `
                -Server $DomainController `
                -Properties DisplayName `
                -ErrorAction Stop


        $UserCache[$ManagerDN] =
            $ManagerObject


        return Get-UserDisplayName `
            -User $ManagerObject
    }
    catch {

        # Preserve DN if manager cannot be resolved.

        return $ManagerDN
    }
}


# ============================================================
# BUILD ACTUAL AD HIERARCHY
# ============================================================

function Get-ActualReportingHierarchy {

    param (
        [Parameter(Mandatory)]
        [object]$User,

        [int]$ADLevel = 0,

        [string[]]$ADReportingPath = @(),

        [string]$VisibleManager = "",

        [int]$VisibleLevel = 0,

        [int]$CollapsedManagerCount = 0,

        [bool]$IsRoot = $false
    )


    # --------------------------------------------------------
    # Circular / Duplicate Protection
    # --------------------------------------------------------

    if (
        -not $VisitedUsers.Add(
            $User.DistinguishedName
        )
    ) {

        $script:Stats.CircularReferences++


        Write-Warning (
            "Circular or duplicate reporting relationship " +
            "detected for '$(
                Get-UserDisplayName -User $User
            )' [$($User.SamAccountName)]. " +
            "Branch stopped."
        )


        return $null
    }


    $script:Stats.UsersDiscovered++


    # --------------------------------------------------------
    # Current User Information
    # --------------------------------------------------------

    $CurrentDisplayName =
        Get-UserDisplayName `
            -User $User


    $CurrentReportingPath = @(
        $ADReportingPath +
        $CurrentDisplayName
    )


    $ReportingPath =
        $CurrentReportingPath -join " > "


    $CurrentADManager =
        Get-ManagerDisplayName `
            -ManagerDN $User.Manager


    # --------------------------------------------------------
    # Get Direct Reports
    # --------------------------------------------------------

    $DirectReports = @(
        Get-DirectReports `
            -ManagerDN $User.DistinguishedName
    )


    # --------------------------------------------------------
    # Build Hierarchy Node
    # --------------------------------------------------------

    $Node = [PSCustomObject]@{

        DistinguishedName      = $User.DistinguishedName

        User                   = $User

        ADLevel                = $ADLevel

        VisibleLevel           = $VisibleLevel

        VisibleManager         = $VisibleManager

        ADManager              = $CurrentADManager

        ReportingPath          = $ReportingPath

        ReportingPathArray     = $CurrentReportingPath

        CollapsedManagerCount  = $CollapsedManagerCount

        IsRoot                 = $IsRoot

        Children               =
            [System.Collections.Generic.List[object]]::new()
    }


    # --------------------------------------------------------
    # Determine Child Context
    # --------------------------------------------------------

    $IncludeCurrentUser =
        $User.Enabled -or $IsRoot


    if ($IncludeCurrentUser) {

        $ChildVisibleManager =
            $CurrentDisplayName


        $ChildVisibleLevel =
            $VisibleLevel + 1


        $ChildCollapsedManagerCount =
            0
    }
    else {

        # Disabled account is collapsed.

        $ChildVisibleManager =
            $VisibleManager


        $ChildVisibleLevel =
            $VisibleLevel


        $ChildCollapsedManagerCount =
            $CollapsedManagerCount + 1
    }


    # --------------------------------------------------------
    # Recursively Process Direct Reports
    # --------------------------------------------------------

    foreach ($DirectReport in $DirectReports) {

        $ChildNode =
            Get-ActualReportingHierarchy `
                -User $DirectReport `
                -ADLevel ($ADLevel + 1) `
                -ADReportingPath $CurrentReportingPath `
                -VisibleManager $ChildVisibleManager `
                -VisibleLevel $ChildVisibleLevel `
                -CollapsedManagerCount $ChildCollapsedManagerCount `
                -IsRoot $false


        if ($null -ne $ChildNode) {

            $Node.Children.Add(
                $ChildNode
            )
        }
    }


    return $Node
}


# ============================================================
# BUILD HIERARCHY
# ============================================================

Write-Host ""
Write-Host "Discovering reporting hierarchy via LDAP..."
Write-Host ""


# Attempt one-query recursive branch discovery first. If unsupported or
# unsuccessful, Get-DirectReports automatically uses the original
# per-manager LDAP traversal.
$null = Initialize-RecursiveReportingDiscovery `
    -RootUser $RootUser


$HierarchyStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$RootNode =
    Get-ActualReportingHierarchy `
        -User $RootUser `
        -ADLevel 0 `
        -ADReportingPath @() `
        -VisibleManager "" `
        -VisibleLevel 0 `
        -CollapsedManagerCount 0 `
        -IsRoot $true

$HierarchyStopwatch.Stop()
$Stats.HierarchyMs = $HierarchyStopwatch.Elapsed.TotalMilliseconds


if ($null -eq $RootNode) {

    Write-Error (
        "Unable to construct reporting hierarchy."
    )

    return
}


# ============================================================
# BATCH ENRICH DISCOVERED USERS
# ============================================================

Initialize-BatchedEnrichment -RootNode $RootNode


# ============================================================
# GET VISIBLE CHILDREN
# ============================================================
#
# Disabled accounts are removed from the visible hierarchy,
# but their enabled descendants are promoted upward.
#
# This is done BEFORE console rendering so the tree connectors
# accurately represent the hierarchy the user actually sees.

function Get-VisibleChildren {

    param (
        [Parameter(Mandatory)]
        [object]$Node
    )


    $VisibleChildren =
        [System.Collections.Generic.List[object]]::new()


    foreach ($Child in $Node.Children) {

        if ($Child.User.Enabled) {

            $VisibleChildren.Add(
                $Child
            )
        }
        else {

            # Disabled child is hidden.
            # Promote its visible descendants.

            $PromotedChildren = @(
                Get-VisibleChildren `
                    -Node $Child
            )


            foreach (
                $PromotedChild in
                $PromotedChildren
            ) {

                $VisibleChildren.Add(
                    $PromotedChild
                )
            }
        }
    }


    return @(
        $VisibleChildren |
            Sort-Object `
                @{ Expression = {
                    Get-UserDisplayName `
                        -User $_.User
                } },
                @{ Expression = {
                    $_.User.SamAccountName
                } }
    )
}


# ============================================================
# ADD MAIN REPORT USER
# ============================================================

function Add-ReportingUser {

    param (
        [Parameter(Mandatory)]
        [object]$Node
    )

    # --------------------------------------------------------
    # Direct Enriched Cache Lookup
    # --------------------------------------------------------
    #
    # Batch enrichment has already populated EnrichedUserCache
    # before report construction begins. Access the dictionary
    # directly on this hot path instead of calling the general
    # Get-EnrichedUser helper for every report row.
    #
    # The helper remains as a safety fallback if a cache entry
    # is unexpectedly missing.

    $CacheLookupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if ($NormalizedUserCache.ContainsKey($Node.DistinguishedName)) {
        $User = $NormalizedUserCache[$Node.DistinguishedName]
        $script:Stats.CacheHits++
    }
    else {
        $script:Stats.ReportCacheFallbacks++
        $RawUser = Get-EnrichedUser -DistinguishedName $Node.DistinguishedName
        if ($null -ne $RawUser) {
            $NormalizationStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $User = ConvertTo-NormalizedUser -User $RawUser
            $NormalizationStopwatch.Stop()
            $script:Stats.FallbackNormalizationMs += $NormalizationStopwatch.Elapsed.TotalMilliseconds
            $script:Stats.UsersNormalized++
            $NormalizedUserCache[$Node.DistinguishedName] = $User
        }
        else {
            $User = $null
        }
    }

    $CacheLookupStopwatch.Stop()
    $script:Stats.ReportCacheLookupMs += $CacheLookupStopwatch.Elapsed.TotalMilliseconds

    if ($null -eq $User) {
        Write-Warning ("Skipping report row for: " + $Node.DistinguishedName)
        return
    }

    # --------------------------------------------------------
    # Prepare Display / Hierarchy Values
    # --------------------------------------------------------

    $DisplayPrepStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if (-not [string]::IsNullOrWhiteSpace($User.DisplayName)) {
        $CurrentDisplayName = $User.DisplayName
    }
    elseif (-not [string]::IsNullOrWhiteSpace($User.SamAccountName)) {
        $CurrentDisplayName = $User.SamAccountName
    }
    else {
        $CurrentDisplayName = $User.DistinguishedName
    }

    $HierarchyName = ("    " * $Node.VisibleLevel) + $CurrentDisplayName

    $DisplayPrepStopwatch.Stop()
    $script:Stats.ReportDisplayPrepMs += $DisplayPrepStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Construct Report Object
    # --------------------------------------------------------

    $PSObjectStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $ReportRow = [PSCustomObject][ordered]@{
        ADLevel               = $Node.ADLevel
        VisibleLevel          = $Node.VisibleLevel
        CollapsedManagerCount = $Node.CollapsedManagerCount
        ReportingPath         = $Node.ReportingPath
        HierarchyName         = $HierarchyName
        DisplayName           = $CurrentDisplayName
        UserPrincipalName     = $User.UserPrincipalName
        SamAccountName        = $User.SamAccountName
        Mail                  = $User.Mail
        Title                 = $User.Title
        Department            = $User.Department
        EmployeeID            = $User.EmployeeID
        Enabled               = $User.Enabled
        VisibleManager        = $Node.VisibleManager
        ADManager             = $Node.ADManager
        DistinguishedName     = $User.DistinguishedName
        extensionAttribute1   = $User.extensionAttribute1
        extensionAttribute2   = $User.extensionAttribute2
        extensionAttribute3   = $User.extensionAttribute3
        extensionAttribute4   = $User.extensionAttribute4
        extensionAttribute5   = $User.extensionAttribute5
        extensionAttribute6   = $User.extensionAttribute6
        extensionAttribute7   = $User.extensionAttribute7
        extensionAttribute8   = $User.extensionAttribute8
        extensionAttribute9   = $User.extensionAttribute9
        extensionAttribute10  = $User.extensionAttribute10
        extensionAttribute11  = $User.extensionAttribute11
        extensionAttribute12  = $User.extensionAttribute12
        extensionAttribute13  = $User.extensionAttribute13
        extensionAttribute14  = $User.extensionAttribute14
        extensionAttribute15  = $User.extensionAttribute15
    }

    $PSObjectStopwatch.Stop()
    $script:Stats.ReportPSObjectMs += $PSObjectStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Add Row to Report Collection
    # --------------------------------------------------------

    $ListAddStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ReportingTree.Add($ReportRow)
    $ListAddStopwatch.Stop()
    $script:Stats.ReportListAddMs += $ListAddStopwatch.Elapsed.TotalMilliseconds
}


# ============================================================
# ADD DISABLED USER
# ============================================================

function Add-DisabledUser {

    param (
        [Parameter(Mandatory)]
        [object]$Node
    )


    if ($NormalizedUserCache.ContainsKey($Node.DistinguishedName)) {
        $User = $NormalizedUserCache[$Node.DistinguishedName]
    }
    else {
        $RawUser = Get-EnrichedUser -DistinguishedName $Node.DistinguishedName
        if ($null -ne $RawUser) {
            $NormalizationStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $User = ConvertTo-NormalizedUser -User $RawUser
            $NormalizationStopwatch.Stop()
            $script:Stats.FallbackNormalizationMs += $NormalizationStopwatch.Elapsed.TotalMilliseconds
            $script:Stats.UsersNormalized++
            $NormalizedUserCache[$Node.DistinguishedName] = $User
        }
        else {
            $User = $null
        }
    }


    if ($null -eq $User) {

        return
    }


    $CurrentDisplayName =
        Get-UserDisplayName `
            -User $User


    $DisabledUsers.Add(

        [PSCustomObject]@{

            ADLevel =
                $Node.ADLevel

            CollapsedManagerCount =
                ($Node.CollapsedManagerCount + 1)

            ReportingPath =
                $Node.ReportingPath

            DisplayName =
                $CurrentDisplayName

            UserPrincipalName =
                $User.UserPrincipalName

            SamAccountName =
                $User.SamAccountName

            Mail =
                $User.Mail

            Title =
                $User.Title

            Department =
                $User.Department

            EmployeeID =
                $User.EmployeeID

            Enabled =
                $User.Enabled

            ADManager =
                $Node.ADManager

            VisibleManager =
                $Node.VisibleManager

            DistinguishedName =
                $User.DistinguishedName

            extensionAttribute1 =
                $User.extensionAttribute1

            extensionAttribute2 =
                $User.extensionAttribute2

            extensionAttribute3 =
                $User.extensionAttribute3

            extensionAttribute4 =
                $User.extensionAttribute4

            extensionAttribute5 =
                $User.extensionAttribute5

            extensionAttribute6 =
                $User.extensionAttribute6

            extensionAttribute7 =
                $User.extensionAttribute7

            extensionAttribute8 =
                $User.extensionAttribute8

            extensionAttribute9 =
                $User.extensionAttribute9

            extensionAttribute10 =
                $User.extensionAttribute10

            extensionAttribute11 =
                $User.extensionAttribute11

            extensionAttribute12 =
                $User.extensionAttribute12

            extensionAttribute13 =
                $User.extensionAttribute13

            extensionAttribute14 =
                $User.extensionAttribute14

            extensionAttribute15 =
                $User.extensionAttribute15
        }
    )
}


# ============================================================
# PROCESS DISABLED USERS
# ============================================================
#
# Disabled users are not rendered in the console tree, so
# process them separately for the disabled-account CSV.

function Find-DisabledUsers {

    param (
        [Parameter(Mandatory)]
        [object]$Node
    )


    if (-not $Node.User.Enabled) {

        Add-DisabledUser `
            -Node $Node
    }


    foreach ($Child in $Node.Children) {

        Find-DisabledUsers `
            -Node $Child
    }
}


$DisabledReportStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Find-DisabledUsers `
    -Node $RootNode
$DisabledReportStopwatch.Stop()
$Stats.DisabledReportMs = $DisabledReportStopwatch.Elapsed.TotalMilliseconds


# ============================================================
# RENDER VISIBLE REPORTING TREE
# ============================================================

function Show-VisibleReportingTree {

    param (
        [Parameter(Mandatory)]
        [object]$Node,

        [string]$Prefix = "",

        [bool]$IsLast = $true,

        [bool]$IsRoot = $false
    )

    $User = $Node.User
    $CurrentDisplayName = Get-UserDisplayName -User $User

    if ($IsRoot) {
        $Connector = ""
    }
    elseif ($IsLast) {
        $Connector = "└── "
    }
    else {
        $Connector = "├── "
    }

    # --------------------------------------------------------
    # Build Console Line (buffered; no Write-Host here)
    # --------------------------------------------------------

    $ConsoleLineStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if ($IsRoot -and -not $User.Enabled) {
        $ConsoleTreeLines.Add(
            "$Prefix$Connector$CurrentDisplayName [$($User.SamAccountName)] [DISABLED]"
        )
    }
    else {
        $ConsoleTreeLines.Add(
            "$Prefix$Connector$CurrentDisplayName [$($User.SamAccountName)]"
        )
    }

    $ConsoleLineStopwatch.Stop()
    $script:Stats.ConsoleBuildMs += $ConsoleLineStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Build Main Report Row
    # --------------------------------------------------------

    $ReportObjectStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Add-ReportingUser -Node $Node
    $ReportObjectStopwatch.Stop()
    $script:Stats.ReportObjectMs += $ReportObjectStopwatch.Elapsed.TotalMilliseconds

    # --------------------------------------------------------
    # Determine Visible Children
    # --------------------------------------------------------

    $VisibleTreeStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $VisibleChildren = @(Get-VisibleChildren -Node $Node)
    $VisibleTreeStopwatch.Stop()
    $script:Stats.VisibleTreeMs += $VisibleTreeStopwatch.Elapsed.TotalMilliseconds

    if ($IsRoot) {
        $ChildPrefix = ""
    }
    elseif ($IsLast) {
        $ChildPrefix = "$Prefix    "
    }
    else {
        $ChildPrefix = "$Prefix│   "
    }

    for ($i = 0; $i -lt $VisibleChildren.Count; $i++) {
        $Child = $VisibleChildren[$i]
        $ChildIsLast = ($i -eq ($VisibleChildren.Count - 1))

        Show-VisibleReportingTree `
            -Node $Child `
            -Prefix $ChildPrefix `
            -IsLast $ChildIsLast `
            -IsRoot $false
    }
}


# ============================================================
# BUILD AND DISPLAY TREE
# ============================================================

# Step 4 optimization:
#   1. Build console lines and CSV report objects in memory.
#   2. Write the entire console tree in a single Write-Host call.
# This preserves the existing visual output while avoiding a host/UI
# round trip for every individual user.

$RenderStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Show-VisibleReportingTree `
    -Node $RootNode `
    -Prefix "" `
    -IsLast $true `
    -IsRoot $true

$RenderStopwatch.Stop()
$Stats.RenderReportMs = $RenderStopwatch.Elapsed.TotalMilliseconds

Write-Host ""
Write-Host "Reporting Tree"
Write-Host "=============="
Write-Host ""

$ConsoleWriteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

if ($ConsoleTreeLines.Count -gt 0) {
    Write-Host ($ConsoleTreeLines -join [Environment]::NewLine)
}

$ConsoleWriteStopwatch.Stop()
$Stats.ConsoleWriteMs = $ConsoleWriteStopwatch.Elapsed.TotalMilliseconds


# ============================================================
# EXPORT MAIN REPORT
# ============================================================

$CsvStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {

    $ReportingTree |
        Export-Csv `
            -Path $ReportingCsvPath `
            -NoTypeInformation `
            -Encoding UTF8 `
            -ErrorAction Stop
}
catch {

    Write-Warning (
        "Unable to export reporting tree to " +
        "'$ReportingCsvPath'. Error: " +
        $_.Exception.Message
    )
}


# ============================================================
# EXPORT DISABLED ACCOUNT REPORT
# ============================================================

if ($DisabledUsers.Count -gt 0) {

    try {

        $DisabledUsers |
            Sort-Object `
                ADLevel,
                DisplayName |
            Export-Csv `
                -Path $DisabledCsvPath `
                -NoTypeInformation `
                -Encoding UTF8 `
                -ErrorAction Stop
    }
    catch {

        Write-Warning (
            "Unable to export disabled account report to " +
            "'$DisabledCsvPath'. Error: " +
            $_.Exception.Message
        )
    }
}
else {

    # Remove stale report from an earlier run.

    if (
        Test-Path $DisabledCsvPath
    ) {

        try {

            Remove-Item `
                -Path $DisabledCsvPath `
                -Force `
                -ErrorAction Stop
        }
        catch {

            Write-Warning (
                "No disabled accounts were found, but " +
                "the previous disabled-account report " +
                "could not be removed."
            )
        }
    }
}


$CsvStopwatch.Stop()
$Stats.CsvExportMs = $CsvStopwatch.Elapsed.TotalMilliseconds


# ============================================================
# QUALITY CONTROL
# ============================================================

$UsersWithCollapsedManagers = @(

    $ReportingTree |

        Where-Object {

            $_.CollapsedManagerCount -gt 0
        }
)


# ============================================================
# FINAL SUMMARY
# ============================================================

$TotalStopwatch.Stop()

Write-Host ""
Write-Host "Report Complete"
Write-Host "==============="
Write-Host ""
Write-Host ("Domain controller   : " + $DomainController)


Write-Host (
    "Users discovered    : " +
    $Stats.UsersDiscovered
)

Write-Host (
    "Users in report     : " +
    $ReportingTree.Count
)

Write-Host (
    "Disabled users      : " +
    $DisabledUsers.Count
)

Write-Host (
    "LDAP discovery calls: " +
    $Stats.LDAPQueries
)

Write-Host (
    "Recursive LDAP users: " +
    $Stats.RecursiveDiscoveryUsers
)

Write-Host (
    "Discovery fallbacks : " +
    $Stats.RecursiveDiscoveryFallbacks
)

Write-Host (
    "Users enriched      : " +
    $Stats.UsersEnriched
)

Write-Host (
    "Enrichment batches  : " +
    $Stats.EnrichmentBatches
)

Write-Host (
    "Enrichment fallbacks: " +
    $Stats.EnrichmentFallbacks
)

Write-Host (
    "Cache hits          : " +
    $Stats.CacheHits
)

Write-Host (
    "Circular references : " +
    $Stats.CircularReferences
)

# ============================================================
# PERFORMANCE SUMMARY
# ============================================================

$AvgLDAPMs = if ($Stats.LDAPQueries -gt 0) {
    $Stats.LDAPBranchMs / $Stats.LDAPQueries
}
else { 0 }

$AvgEnrichmentBatchMs = if ($Stats.EnrichmentBatches -gt 0) {
    $Stats.EnrichmentLdapMs / $Stats.EnrichmentBatches
}
else { 0 }

Write-Host ""
Write-Host "Performance"
Write-Host "==========="
Write-Host ""
Write-Host ("Total elapsed          : {0:hh\:mm\:ss\.fff}" -f $TotalStopwatch.Elapsed)
Write-Host ("DC discovery           : {0:N2} sec" -f ($Stats.DCDiscoveryMs / 1000))
Write-Host ("Manager AD search      : {0:N2} sec" -f ($Stats.ManagerSearchMs / 1000))
Write-Host ("Hierarchy build        : {0:N2} sec" -f ($Stats.HierarchyMs / 1000))
Write-Host ("Discovery LDAP time    : {0:N2} sec" -f ($Stats.LDAPBranchMs / 1000))
Write-Host ("  LDAP discovery calls : {0}" -f $Stats.LDAPQueries)
Write-Host ("  Avg discovery call   : {0:N2} ms" -f $AvgLDAPMs)
Write-Host ("  Recursive mode       : {0}" -f $RecursiveDiscoveryLoaded)
Write-Host ("Raw LDAP enrichment    : {0:N2} sec" -f ($Stats.EnrichmentMs / 1000))
Write-Host ("  Users parsed         : {0}" -f $Stats.UsersParsed)
Write-Host ("  LDAP query time      : {0:N2} sec" -f ($Stats.EnrichmentLdapMs / 1000))
Write-Host ("  LDAP result parsing  : {0:N2} sec" -f ($Stats.EnrichmentParseMs / 1000))
Write-Host ("  Enrichment batches   : {0}" -f $Stats.EnrichmentBatches)
Write-Host ("  Batch size           : {0}" -f $EnrichmentBatchSize)
Write-Host ("  Avg LDAP batch       : {0:N2} ms" -f $AvgEnrichmentBatchMs)
Write-Host ("  Fallbacks            : {0}" -f $Stats.EnrichmentFallbacks)

if ($Stats.EnrichmentFallbacks -gt 0 -or $Stats.FallbackNormalizationMs -gt 0) {
    Write-Host ("Fallback AD handling   : {0:N2} sec" -f ($Stats.FallbackNormalizationMs / 1000))
    Write-Host ("  Fallback users norm. : {0}" -f $Stats.UsersNormalized)
}

Write-Host ("Report/tree build      : {0:N2} sec" -f ($Stats.RenderReportMs / 1000))
Write-Host ("Disabled report build  : {0:N2} sec" -f ($Stats.DisabledReportMs / 1000))
Write-Host ("Console write          : {0:N2} sec" -f ($Stats.ConsoleWriteMs / 1000))
Write-Host ("CSV export             : {0:N2} sec" -f ($Stats.CsvExportMs / 1000))


# ============================================================
# OUTPUT LOCATIONS
# ============================================================

Write-Host ""
Write-Host "Reporting tree CSV:"
Write-Host $ReportingCsvPath


if ($DisabledUsers.Count -gt 0) {

    Write-Host ""
    Write-Host "Disabled accounts CSV:"
    Write-Host $DisabledCsvPath
}


# ============================================================
# QC SUMMARY
# ============================================================

Write-Host ""
Write-Host "Quality Control"
Write-Host "==============="
Write-Host ""


Write-Host (
    "Users with collapsed disabled managers: " +
    $UsersWithCollapsedManagers.Count
)


if (
    $UsersWithCollapsedManagers.Count -gt 0
) {

    Write-Host ""
    Write-Host (
        "Users affected by disabled managers:"
    )

    Write-Host ""


    $UsersWithCollapsedManagers |

        Select-Object `
            DisplayName,
            UserPrincipalName,
            ADManager,
            VisibleManager,
            CollapsedManagerCount,
            ReportingPath |

        Format-Table -AutoSize
}


Write-Host ""
