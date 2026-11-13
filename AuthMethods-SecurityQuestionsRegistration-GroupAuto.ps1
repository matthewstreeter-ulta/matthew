#Requires -Version 7.0
<#
.SYNOPSIS
    Export an Entra group's direct user members and check Security Questions registration.
    AccountEnabled is included in both CSVs. -AccountStatus filters only the
    registration report; membership export always includes all direct user members.
    Unknown account status is retained only when -AccountStatus All is used.
    Console summary includes enabled-user Security Questions registration percentage.
    A separate store-level CSV groups users by EmployeeId and prioritizes follow-up.
.NOTES
    Requires Microsoft.Graph.Authentication and delegated Graph permissions:
    Group.Read.All, User.Read.All, AuditLog.Read.All.
    The signed-in identity must also have the necessary Entra role to read
    authentication method registration reports (for example Reports Reader).
    Optional LDAPS enrichment requires -IncludePasswordExpiration -DomainController <FQDN>.
    Password timestamps are UTC; LDAP failure leaves password status Unknown.
    v2.5.3.2 add multiple OUs search, with separate paging cookies and combined SID indexing.
    v2.5.2 adds SecretManagement credential retrieval with an interactive fallback.
    v2.5.1 adds paging response-control diagnostics and rejects an ambiguous full final page.
    The computed expiry attribute may not be returned for all accounts in a subtree query.
    Group membership is direct membership; use transitiveMembers if nested
    group expansion is desired.
#>
[CmdletBinding()]
param(
    [string]$GroupName = 'GROUPNAME',
    [string]$GroupId,
    [string]$OutputDirectory = '\Path\To\EntraReports',
    [ValidateRange(1,999)] [int]$MemberPageSize = 999,
    [ValidateRange(1,1000)] [int]$ReportPageSize = 1000,
    [ValidateSet('All','EnabledOnly','DisabledOnly')] [string]$AccountStatus = 'All',
    [switch]$SkipMembershipExport,
    [switch]$IncludePasswordExpiration,
    [string]$DomainController,
    [int]$LdapPort = 636,
    [ValidateRange(100,1000)] [int]$LdapPageSize = 500,
    [Alias('LdapSearchBase')] [string[]]$LdapSearchBases = @('OU=Users,OU=Corporate,OU=company,DC=example,DC=com', 'OU=Stores,OU=company,DC=example,DC=com'),
    [pscredential]$AdCredential,
    [ValidateSet('Auto','Stored','Interactive')] [string]$AdCredentialMode = 'Auto',
    [string]$AdSecretName = 'Company-AD-LDAPS',
    [string]$AdSecretVault = 'LocalVault'
)

$ErrorActionPreference = 'Stop'
$overall = [Diagnostics.Stopwatch]::StartNew()
$timings = [ordered]@{}

function Resolve-AdCredential {
    param(
        [ValidateSet('Auto','Stored','Interactive')] [string]$Mode,
        [string]$SecretName,
        [string]$SecretVault
    )

    if ($Mode -ne 'Interactive') {
        try {
            Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
            $vault = Get-SecretVault -Name $SecretVault -ErrorAction Stop
            if ($null -eq $vault) { throw "Secret vault '$SecretVault' is not registered." }
            $secret = Get-Secret -Name $SecretName -Vault $SecretVault -ErrorAction Stop
            if ($secret -isnot [pscredential]) {
                throw "Secret '$SecretName' is not a PSCredential."
            }
            Write-Host ('AD credential loaded from SecretManagement vault: {0}' -f $SecretVault) -ForegroundColor Green
            return $secret
        } catch {
            if ($Mode -eq 'Stored') {
                throw ('Unable to retrieve AD credential from vault {0}: {1}' -f $SecretVault, $_.Exception.Message)
            }
            Write-Warning ('Stored AD credential unavailable ({0}). Switching to interactive authentication.' -f $_.Exception.Message)
        }
    }

    $credential = Get-Credential -Message 'Enter read-only AD credentials for LDAPS'
    if ($null -eq $credential) { throw 'AD credentials were not provided.' }
    return $credential
}

function Get-GraphPagedResults {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Label)
    $items = [Collections.Generic.List[object]]::new()
    $page = 0
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while (-not [string]::IsNullOrWhiteSpace($Uri)) {
        $page++
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        foreach ($item in @($response.value)) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $rate = if ($clock.Elapsed.TotalSeconds -gt 0) { [math]::Round($items.Count / $clock.Elapsed.TotalSeconds) } else { 0 }
        Write-Host ('{0}: {1:N0} records | Page {2:N0} | {3:N0}/sec | Elapsed {4}' -f $Label, $items.Count, $page, $rate, $clock.Elapsed.ToString('hh\:mm\:ss')) -ForegroundColor Yellow
        $Uri = $response.'@odata.nextLink'
    }
    $clock.Stop()
    return ,$items
}

function Get-LdapValue {
    param($Entry, [string]$Name)
    $attr = $Entry.Attributes[$Name]
    if ($null -eq $attr -or $attr.Count -eq 0) { return $null }
    return [string]$attr[0]
}

function Convert-SidBytesToString {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 8) { return $null }
    $subCount = [int]$Bytes[1]
    if ($Bytes.Length -lt (8 + 4 * $subCount)) { return $null }
    [ulong]$authority = 0
    for ($i = 2; $i -lt 8; $i++) { $authority = ($authority -shl 8) -bor [ulong]$Bytes[$i] }
    $parts = [Collections.Generic.List[string]]::new()
    $parts.Add(('S-{0}-{1}' -f $Bytes[0], $authority))
    for ($i = 0; $i -lt $subCount; $i++) {
        $offset = 8 + 4 * $i
        [uint64]$sub = [uint64]$Bytes[$offset] -bor ([uint64]$Bytes[$offset+1] -shl 8) -bor ([uint64]$Bytes[$offset+2] -shl 16) -bor ([uint64]$Bytes[$offset+3] -shl 24)
        $parts.Add([string]$sub)
    }
    return ($parts -join '-')
}

function Get-PasswordDetails {
    param($Entry)
    $lastRaw = Get-LdapValue $Entry 'pwdLastSet'
    $expiryRaw = Get-LdapValue $Entry 'msDS-UserPasswordExpiryTimeComputed'
    $uacRaw = Get-LdapValue $Entry 'userAccountControl'
    $lastSet = $null
    $expiry = $null
    $status = 'Unknown'
    $never = $null
    $days = $null
    try {
        if ($lastRaw -and [long]$lastRaw -gt 0) { $lastSet = [datetime]::FromFileTimeUtc([long]$lastRaw) }
        if ($uacRaw) { $never = (([long]$uacRaw -band 65536) -ne 0) }
        if ($never -eq $true -or $expiryRaw -eq '9223372036854775807') { $status = 'Never' }
        elseif ($expiryRaw -eq '0') { $status = 'Must Change Password' }
        elseif ($expiryRaw) {
            $expiry = [datetime]::FromFileTimeUtc([long]$expiryRaw)
            $days = [math]::Ceiling(($expiry - [datetime]::UtcNow).TotalDays)
            $status = if ($expiry -le [datetime]::UtcNow) { 'Expired' } else { 'Active' }
        }
    } catch { $status = 'Invalid'; $expiry = $null; $days = $null }
    return [pscustomobject]@{
        PasswordLastSet = if ($lastSet) { $lastSet.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC' } else { $null }
        PasswordExpirationDate = if ($expiry) { $expiry.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC' } else { $null }
        DaysUntilExpiration = $days
        PasswordNeverExpires = $never
        PasswordStatus = $status
    }
}

function Get-AdPasswordBySid {
    param([string]$Server, [int]$Port, [int]$PageSize, [string[]]$Bases, [pscredential]$Credential)
    Add-Type -AssemblyName System.DirectoryServices.Protocols
    $identifier = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]::new($Server, $Port, $false, $false)
    $connection = [System.DirectoryServices.Protocols.LdapConnection]::new($identifier)
    $lookup = @{}
    $totalPages = 0
    $totalEntries = 0
    $duplicateSidCount = 0
    $missingSidCount = 0
    $baseMetrics = [System.Collections.Generic.List[object]]::new()
    try {
        $connection.SessionOptions.SecureSocketLayer = $true
        $connection.SessionOptions.ProtocolVersion = 3
        $connection.AuthType = [System.DirectoryServices.Protocols.AuthType]::Basic
        $connection.Credential = [System.Net.NetworkCredential]::new($Credential.UserName, $Credential.GetNetworkCredential().Password)
        $connection.Bind()
        [string[]]$attrs = @('objectSid', 'pwdLastSet', 'msDS-UserPasswordExpiryTimeComputed', 'userAccountControl')
        if ($null -eq $Bases -or $Bases.Count -eq 0) { throw 'At least one LDAP search base must be specified.' }
        foreach ($base in $Bases) {
            if ([string]::IsNullOrWhiteSpace($base)) { throw 'LDAP search base cannot be blank.' }
            Write-Host ('LDAPS connected | Search base: {0}' -f $base) -ForegroundColor Green
            $cookie = [byte[]]@()
            $pageNumber = 0
            $baseEntries = 0
            $baseUnique = 0
            do {
                $request = [System.DirectoryServices.Protocols.SearchRequest]::new($base, '(&(objectCategory=person)(objectClass=user))', [System.DirectoryServices.Protocols.SearchScope]::Subtree, $attrs)
                $page = [System.DirectoryServices.Protocols.PageResultRequestControl]::new($PageSize)
                $page.Cookie = $cookie
                [void]$request.Controls.Add($page)
                $response = $connection.SendRequest($request)
                $pageNumber++
                $totalPages++
                $baseEntries += $response.Entries.Count
                $totalEntries += $response.Entries.Count
                foreach ($entry in $response.Entries) {
                    $sidAttribute = $entry.Attributes['objectSid']
                    if ($null -eq $sidAttribute -or $sidAttribute.Count -eq 0) { $missingSidCount++; continue }
                    $sidBytes = [byte[]]$sidAttribute.GetValues([byte[]])[0]
                    $sid = Convert-SidBytesToString $sidBytes
                    if ($sid) {
                        if ($lookup.ContainsKey($sid)) { $duplicateSidCount++ }
                        else { $lookup[$sid] = Get-PasswordDetails $entry; $baseUnique++ }
                    } else { $missingSidCount++ }
                }
                $pageResponse = $null
                foreach ($control in $response.Controls) {
                    if ($control -is [System.DirectoryServices.Protocols.PageResultResponseControl]) {
                        $pageResponse = $control
                        break
                    }
                }
                if ($null -eq $pageResponse) { throw ('No LDAP paging response control for base: {0}' -f $base) }
                $cookie = $pageResponse.Cookie
                $cookieLength = if ($null -eq $cookie) { 0 } else { $cookie.Length }
                Write-Host ('LDAP page: {0:N0} | Entries: {1:N0} | Indexed (OU): {2:N0} | Cookie bytes: {3:N0} | Result: {4}' -f $pageNumber, $response.Entries.Count, $baseEntries, $cookieLength, $response.ResultCode) -ForegroundColor Yellow
                if ($cookieLength -eq 0 -and $response.Entries.Count -ge $PageSize) {
                    throw ('LDAP paging ended after a full page without continuation cookie for base: {0}' -f $base)
                }
            } while ($null -ne $cookie -and $cookie.Length -gt 0)
            $baseMetrics.Add([pscustomobject]@{ SearchBase = $base; Pages = $pageNumber; Entries = $baseEntries; UniqueSidsAdded = $baseUnique })
            Write-Host ('OU complete | Pages: {0:N0} | Entries: {1:N0} | New unique SIDs: {2:N0}' -f $pageNumber, $baseEntries, $baseUnique) -ForegroundColor Green
        }
        Write-Host ('LDAP complete | Pages: {0:N0} | Entries: {1:N0} | Unique SIDs: {2:N0} | Duplicate SIDs: {3:N0} | Missing/invalid SIDs: {4:N0}' -f $totalPages, $totalEntries, $lookup.Count, $duplicateSidCount, $missingSidCount) -ForegroundColor Green
        return [pscustomobject]@{
            Lookup = $lookup
            Pages = $totalPages
            Entries = $totalEntries
            DuplicateSids = $duplicateSidCount
            MissingSids = $missingSidCount
            SearchBases = $Bases
            BaseMetrics = @($baseMetrics.ToArray())
        }
    } finally { $connection.Dispose() }
}

try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication is not installed. Run Install-Module Microsoft.Graph.Authentication -Scope CurrentUser.'
    }
    Import-Module Microsoft.Graph.Authentication
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes @('Group.Read.All','User.Read.All','AuditLog.Read.All') -NoWelcome
    }
    $ctx = Get-MgContext
    $required = @('Group.Read.All','User.Read.All','AuditLog.Read.All')
    if ($ctx.AuthType -eq 'Delegated') {
        $missing = @($required | Where-Object { $_ -notin @($ctx.Scopes) })
        if ($missing.Count -gt 0) {
            throw ('Current Graph connection lacks scopes: {0}. Reconnect with Connect-MgGraph -Scopes Group.Read.All,User.Read.All,AuditLog.Read.All.' -f ($missing -join ', '))
        }
    }

    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $safeName = $GroupName -replace '[\\/:*?"<>|]', '_'
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = if ($GroupId) { $GroupId } else { 'EntraGroup' } }
    $membersCsv = Join-Path $OutputDirectory ($safeName + '-Members.csv')
    $filterSuffix = if ($AccountStatus -eq 'All') { '' } else { '-' + $AccountStatus }
    $outputCsv = Join-Path $OutputDirectory ($safeName + '-SecurityQuestionRegistration' + $filterSuffix + '.csv')
    $storeCsv = Join-Path $OutputDirectory ($safeName + '-SecurityQuestionsByStore' + $filterSuffix + '.csv')

    Write-Host "`nResolving Entra group..." -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    if ($GroupId) {
        $group = Invoke-MgGraphRequest -Method GET -Uri ('https://graph.microsoft.com/v1.0/groups/{0}?$select=id,displayName' -f [uri]::EscapeDataString($GroupId)) -OutputType PSObject
    } else {
        $escaped = $GroupName.Replace("'", "''")
        $filter = [uri]::EscapeDataString("displayName eq '$escaped'")
        $matches = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName&`$top=100") -OutputType PSObject
        if (@($matches.value).Count -eq 0) { throw "No Entra group found with displayName '$GroupName'." }
        if (@($matches.value).Count -gt 1) { throw "Multiple groups share displayName '$GroupName'. Rerun with -GroupId to choose the intended group." }
        $group = @($matches.value)[0]
    }
    $watch.Stop(); $timings['Group resolution'] = $watch.Elapsed
    Write-Host ('Group: {0} | Id: {1}' -f $group.displayName, $group.id) -ForegroundColor Green

    Write-Host "`nRetrieving direct user members..." -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $memberUri = 'https://graph.microsoft.com/v1.0/groups/{0}/members/microsoft.graph.user?$select=id,displayName,userPrincipalName,mail,employeeId,accountEnabled,onPremisesSecurityIdentifier&$top={1}' -f $group.id, $MemberPageSize
    $memberRecords = Get-GraphPagedResults -Uri $memberUri -Label 'Group members'
    $members = @($memberRecords | ForEach-Object {
        [pscustomobject]@{
            Id = $_.id
            userPrincipalName = $_.userPrincipalName
            displayName = $_.displayName
            mail = $_.mail
            employeeId = $_.employeeId
            AccountEnabled = if ($null -eq $_.accountEnabled) { $null } else { [bool]$_.accountEnabled }
            OnPremisesSecurityIdentifier = $_.onPremisesSecurityIdentifier
        }
    })
    $watch.Stop(); $timings['Group retrieval'] = $watch.Elapsed
    if ($members.Count -eq 0) { throw 'Group has no directly assigned user members; report not generated.' }
    Write-Host ('Users retrieved: {0:N0}' -f $members.Count) -ForegroundColor Green

    $enabledMembers = @($members | Where-Object { $null -ne $_.AccountEnabled -and $_.AccountEnabled -eq $true }).Count
    $disabledMembers = @($members | Where-Object { $null -ne $_.AccountEnabled -and $_.AccountEnabled -eq $false }).Count
    $unknownMembers = $members.Count - $enabledMembers - $disabledMembers
    $statusMessage = 'Account status: Enabled {0:N0} | Disabled {1:N0}' -f $enabledMembers, $disabledMembers
    if ($unknownMembers -gt 0) { $statusMessage += ' | Unknown {0:N0}' -f $unknownMembers }
    Write-Host $statusMessage

    if (-not $SkipMembershipExport) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $members | Export-Csv -Path $membersCsv -Encoding utf8
        $watch.Stop(); $timings['Membership CSV export'] = $watch.Elapsed
        Write-Host "Membership CSV: $membersCsv"
    }

    $adPasswords = @{}
    $ldapAvailable = $false
    $ldapMetrics = $null
    if ($IncludePasswordExpiration) {
        if ([string]::IsNullOrWhiteSpace($DomainController)) { throw '-DomainController is required with -IncludePasswordExpiration.' }
        if ($null -eq $AdCredential) { $AdCredential = Resolve-AdCredential -Mode $AdCredentialMode -SecretName $AdSecretName -SecretVault $AdSecretVault }
        if ($null -eq $AdCredential) { throw 'AD credentials were not provided.' }
        Write-Host "`nRetrieving AD password expiration information over LDAPS..." -ForegroundColor Cyan
        $watch = [Diagnostics.Stopwatch]::StartNew()
        try {
            $ldapMetrics = Get-AdPasswordBySid -Server $DomainController -Port $LdapPort -PageSize $LdapPageSize -Bases $LdapSearchBases -Credential $AdCredential
            $adPasswords = $ldapMetrics.Lookup
            $ldapAvailable = $true
        } catch {
            Write-Warning ('LDAPS enrichment unavailable: {0}. Password fields will be Unknown; Graph reporting will continue.' -f $_.Exception.Message)
        } finally {
            $watch.Stop()
            $timings['LDAPS password enrichment'] = $watch.Elapsed
            $AdCredential = $null
        }
    }

    $selectedMembers = @(switch ($AccountStatus) {
        'EnabledOnly' { $members | Where-Object { $null -ne $_.AccountEnabled -and $_.AccountEnabled -eq $true } }
        'DisabledOnly' { $members | Where-Object { $null -ne $_.AccountEnabled -and $_.AccountEnabled -eq $false } }
        default { $members }
    })
    Write-Host ('Account filter: {0} | Users to process: {1:N0}' -f $AccountStatus, $selectedMembers.Count) -ForegroundColor Cyan
    if ($selectedMembers.Count -eq 0) { Write-Warning 'No users match the selected account-status filter. No registration report generated.'; return }

    Write-Host "`nRetrieving authentication method registration report..." -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $reportUri = 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top={0}' -f $ReportPageSize
    $registrations = Get-GraphPagedResults -Uri $reportUri -Label 'Registration report'
    $watch.Stop(); $timings['Graph registration retrieval'] = $watch.Elapsed

    $watch = [Diagnostics.Stopwatch]::StartNew()
    $byId = @{}
    foreach ($reg in $registrations) {
        if (-not [string]::IsNullOrWhiteSpace([string]$reg.id)) { $byId[[string]$reg.id] = $reg }
    }
    $watch.Stop(); $timings['Lookup build'] = $watch.Elapsed

    Write-Host "`nChecking group members..." -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $matched = 0; $unmatched = 0; $security = 0; $noSecurity = 0; $processed = 0
    $statusTotals = @{
        Enabled = @{ Total = 0; Questions = 0; NoQuestions = 0; NotFound = 0 }
        Disabled = @{ Total = 0; Questions = 0; NoQuestions = 0; NotFound = 0 }
        Unknown = @{ Total = 0; Questions = 0; NoQuestions = 0; NotFound = 0 }
    }
    $results = foreach ($member in $selectedMembers) {
        $statusKey = if ($null -eq $member.AccountEnabled) { 'Unknown' } elseif ($member.AccountEnabled) { 'Enabled' } else { 'Disabled' }
        $statusTotals[$statusKey].Total++
        $processed++
        $reg = $null
        if ($byId.ContainsKey([string]$member.Id)) { $reg = $byId[[string]$member.Id]; $matched++ } else { $unmatched++; $statusTotals[$statusKey].NotFound++ }
        $hasQuestions = $null
        if ($null -ne $reg) {
            $hasQuestions = [bool](@($reg.methodsRegistered) -contains 'securityQuestion')
            if ($hasQuestions) { $security++; $statusTotals[$statusKey].Questions++ } else { $noSecurity++; $statusTotals[$statusKey].NoQuestions++ }
        }
        if ($processed % 1000 -eq 0 -or $processed -eq $selectedMembers.Count) {
            Write-Progress -Activity 'Checking Security Questions registration' -Status ("$processed / $($selectedMembers.Count)") -PercentComplete (100 * $processed / $selectedMembers.Count)
        }
        $password = $null
        if ($IncludePasswordExpiration -and $ldapAvailable -and $member.OnPremisesSecurityIdentifier) {
            $sidKey = [string]$member.OnPremisesSecurityIdentifier
            if ($adPasswords.ContainsKey($sidKey)) { $password = $adPasswords[$sidKey] }
        }
        [pscustomobject]@{
            ObjectId = $member.Id
            DisplayName = $member.displayName
            UserPrincipalName = $member.userPrincipalName
            Mail = $member.mail
            EmployeeId = $member.employeeId
            AccountEnabled = $member.AccountEnabled
            SecurityQuestionsRegistered = $hasQuestions
            MethodsRegistered = if ($null -ne $reg) { @($reg.methodsRegistered) -join '; ' } else { $null }
            RegistrationRecordFound = ($null -ne $reg)
            PasswordLastSet = if ($password) { $password.PasswordLastSet } else { $null }
            PasswordExpirationDate = if ($password) { $password.PasswordExpirationDate } else { $null }
            DaysUntilExpiration = if ($password) { $password.DaysUntilExpiration } else { $null }
            PasswordNeverExpires = if ($password) { $password.PasswordNeverExpires } else { $null }
            PasswordStatus = if ($password) { $password.PasswordStatus } else { 'Unknown' }
        }
    }
    Write-Progress -Activity 'Checking Security Questions registration' -Completed
    $watch.Stop(); $timings['Group comparison'] = $watch.Elapsed

    $watch = [Diagnostics.Stopwatch]::StartNew()
    $results | Export-Csv -Path $outputCsv -Encoding utf8
    $watch.Stop(); $timings['Report CSV export'] = $watch.Elapsed

    # Store-level aggregation uses the same processed population as the user report.
    # Missing registration records remain separate from confirmed No Questions.
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $storeBuckets = @{}
    foreach ($user in $results) {
        $storeId = [string]$user.EmployeeId
        if ([string]::IsNullOrWhiteSpace($storeId)) { $storeId = '(Missing EmployeeId)' } else { $storeId = $storeId.Trim() }
        if (-not $storeBuckets.ContainsKey($storeId)) {
            $storeBuckets[$storeId] = [pscustomobject]@{
                EmployeeId = $storeId
                TotalUsers = 0
                Enabled = 0
                Disabled = 0
                Unknown = 0
                EnabledWithQuestions = 0
                EnabledWithoutQuestions = 0
                EnabledNoRecord = 0
                DisabledWithQuestions = 0
                DisabledWithoutQuestions = 0
                DisabledNoRecord = 0
                UnknownWithQuestions = 0
                UnknownWithoutQuestions = 0
                UnknownNoRecord = 0
                EnabledPasswordExpired = 0
                EnabledPasswordExpiring7Days = 0
                EnabledPasswordExpiring14Days = 0
                EnabledPasswordExpiring30Days = 0
                EnabledPasswordUnknown = 0
            }
        }
        $bucket = $storeBuckets[$storeId]
        $bucket.TotalUsers++
        $status = if ($null -eq $user.AccountEnabled) { 'Unknown' } elseif ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }
        $bucket.$status++
        $suffix = if (-not $user.RegistrationRecordFound) { 'NoRecord' } elseif ($user.SecurityQuestionsRegistered) { 'WithQuestions' } else { 'WithoutQuestions' }
        $propertyName = $status + $suffix
        $bucket.$propertyName++
        if ($status -eq 'Enabled' -and $IncludePasswordExpiration) {
            if ($user.PasswordStatus -eq 'Expired' -or $user.PasswordStatus -eq 'Must Change Password') { $bucket.EnabledPasswordExpired++ }
            elseif ($user.PasswordStatus -eq 'Active' -and $null -ne $user.DaysUntilExpiration) {
                $remaining = [int]$user.DaysUntilExpiration
                if ($remaining -le 7) { $bucket.EnabledPasswordExpiring7Days++ }
                if ($remaining -le 14) { $bucket.EnabledPasswordExpiring14Days++ }
                if ($remaining -le 30) { $bucket.EnabledPasswordExpiring30Days++ }
            }
            elseif ($user.PasswordStatus -in @('Unknown','Invalid')) { $bucket.EnabledPasswordUnknown++ }
        }
    }
    $storeReport = @($storeBuckets.Values | ForEach-Object {
        $bucket = $_
        $needsFollowUp = $bucket.EnabledWithoutQuestions + $bucket.EnabledNoRecord
        [pscustomobject]@{
            EmployeeId = $bucket.EmployeeId
            TotalUsers = $bucket.TotalUsers
            Enabled = $bucket.Enabled
            Disabled = $bucket.Disabled
            Unknown = $bucket.Unknown
            EnabledWithQuestions = $bucket.EnabledWithQuestions
            EnabledWithoutQuestions = $bucket.EnabledWithoutQuestions
            EnabledNoRecord = $bucket.EnabledNoRecord
            EnabledNeedsFollowUp = $needsFollowUp
            EnabledRegistrationPercent = if ($bucket.Enabled -gt 0) { [math]::Round(100.0 * $bucket.EnabledWithQuestions / $bucket.Enabled, 2) } else { $null }
            DisabledWithQuestions = $bucket.DisabledWithQuestions
            DisabledWithoutQuestions = $bucket.DisabledWithoutQuestions
            DisabledNoRecord = $bucket.DisabledNoRecord
            UnknownWithQuestions = $bucket.UnknownWithQuestions
            UnknownWithoutQuestions = $bucket.UnknownWithoutQuestions
            UnknownNoRecord = $bucket.UnknownNoRecord
            EnabledPasswordExpired = $bucket.EnabledPasswordExpired
            EnabledPasswordExpiring7Days = $bucket.EnabledPasswordExpiring7Days
            EnabledPasswordExpiring14Days = $bucket.EnabledPasswordExpiring14Days
            EnabledPasswordExpiring30Days = $bucket.EnabledPasswordExpiring30Days
            EnabledPasswordUnknown = $bucket.EnabledPasswordUnknown
        }
    } | Sort-Object -Property @{ Expression = 'EnabledNeedsFollowUp'; Descending = $true }, @{ Expression = 'Enabled'; Descending = $true }, EmployeeId)
    $storeReport | Export-Csv -Path $storeCsv -Encoding utf8
    $watch.Stop(); $timings['Store report aggregation'] = $watch.Elapsed
    $overall.Stop()
    # Summary counts are for processed users; the group status counts above are for all members.
    $enabledProcessed = $statusTotals.Enabled.Total
    $enabledQuestions = $statusTotals.Enabled.Questions
    $enabledPercent = if ($enabledProcessed -gt 0) { '{0:N2}%' -f (100.0 * $enabledQuestions / $enabledProcessed) } else { 'N/A (no enabled users processed)' }

    function Write-SummaryLine {
        param([string]$Label, [string]$Value)
        Write-Host ('{0,-29}: {1,12}' -f $Label, $Value)
    }
    function Format-Count {
        param([long]$Value)
        return ('{0:N0}' -f $Value)
    }

    Write-Host "`nReport Complete`n=============== " -ForegroundColor Green
    Write-SummaryLine 'Group members' (Format-Count $members.Count)
    Write-SummaryLine 'Account status filter' $AccountStatus
    Write-SummaryLine 'Users processed' (Format-Count $selectedMembers.Count)
    Write-SummaryLine 'Enabled users (group)' (Format-Count $enabledMembers)
    Write-SummaryLine 'Disabled users (group)' (Format-Count $disabledMembers)
    if ($unknownMembers -gt 0) { Write-SummaryLine 'Unknown status (group)' (Format-Count $unknownMembers) }
    Write-SummaryLine 'Registration records' (Format-Count $registrations.Count)
    Write-SummaryLine 'Users matched' (Format-Count $matched)
    Write-SummaryLine 'Users not found' (Format-Count $unmatched)
    Write-SummaryLine 'Security Questions' (Format-Count $security)
    Write-SummaryLine 'No Security Questions' (Format-Count $noSecurity)

    Write-Host "`nSecurity Questions by Account Status (processed users)" -ForegroundColor Cyan
    Write-Host ('{0,-10} {1,10} {2,12} {3,14} {4,11}' -f 'Status', 'Total', 'Questions', 'No Questions', 'No Record')
    Write-Host ('-' * 64)
    $statusesToDisplay = @('Enabled','Disabled')
    if ($statusTotals['Unknown'].Total -gt 0) { $statusesToDisplay += 'Unknown' }
    foreach ($statusKey in $statusesToDisplay) {
        $t = $statusTotals[$statusKey]
        Write-Host ('{0,-10} {1,10:N0} {2,12:N0} {3,14:N0} {4,11:N0}' -f $statusKey, $t.Total, $t.Questions, $t.NoQuestions, $t.NotFound)
    }

    Write-Host "`nEnabled Users with Security Questions (processed users)" -ForegroundColor Cyan
    Write-SummaryLine 'Registered enabled users' (Format-Count $enabledQuestions)
    Write-SummaryLine 'Total enabled users' (Format-Count $enabledProcessed)
    Write-SummaryLine 'Registration percentage' $enabledPercent
    if ($AccountStatus -eq 'DisabledOnly') {
        Write-Host 'Note: Enabled-user percentage is unavailable with -AccountStatus DisabledOnly.' -ForegroundColor DarkYellow
    }

    Write-Host "`nStore-Level Follow-Up (top 15 by enabled users needing follow-up)" -ForegroundColor Cyan
    Write-Host ('{0,-22} {1,9} {2,9} {3,11} {4,11} {5,10}' -f 'EmployeeId', 'Enabled', 'Disabled', 'Questions', 'Follow-up', 'Rate')
    Write-Host ('-' * 78)
    foreach ($store in @($storeReport | Select-Object -First 15)) {
        $rate = if ($null -eq $store.EnabledRegistrationPercent) { 'N/A' } else { '{0:N2}%' -f $store.EnabledRegistrationPercent }
        Write-Host ('{0,-22} {1,9:N0} {2,9:N0} {3,11:N0} {4,11:N0} {5,10}' -f $store.EmployeeId, $store.Enabled, $store.Disabled, $store.EnabledWithQuestions, $store.EnabledNeedsFollowUp, $rate)
    }
    Write-Host ('Store/EmployeeId groups: {0:N0} | Full breakdown: {1}' -f $storeReport.Count, $storeCsv)
    if ($AccountStatus -ne 'All') { Write-Host ('Note: Store counts reflect the {0} filter, not all group members.' -f $AccountStatus) -ForegroundColor DarkYellow }

    if ($IncludePasswordExpiration) {
        $sidEligible = @($selectedMembers | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.OnPremisesSecurityIdentifier) })
        $sidMatched = @($sidEligible | Where-Object { $ldapAvailable -and $adPasswords.ContainsKey([string]$_.OnPremisesSecurityIdentifier) }).Count
        $sidUnmatched = $sidEligible.Count - $sidMatched
        $noSid = $selectedMembers.Count - $sidEligible.Count
        $coveragePercent = if ($sidEligible.Count -gt 0) { '{0:N2}%' -f (100.0 * $sidMatched / $sidEligible.Count) } else { 'N/A' }
        Write-Host "`nAD Password Enrichment Coverage (processed users)" -ForegroundColor Cyan
        Write-SummaryLine 'LDAP search bases' ($LdapSearchBases -join ' | ')
        Write-SummaryLine 'LDAP retrieval status' $(if ($ldapAvailable) { 'Complete' } else { 'Incomplete' })
        if ($null -ne $ldapMetrics) {
            foreach ($baseMetric in $ldapMetrics.BaseMetrics) {
                Write-Host ('  OU: {0}' -f $baseMetric.SearchBase)
                Write-SummaryLine '    LDAP pages' (Format-Count $baseMetric.Pages)
                Write-SummaryLine '    AD entries' (Format-Count $baseMetric.Entries)
                Write-SummaryLine '    Unique SIDs added' (Format-Count $baseMetric.UniqueSidsAdded)
            }
            Write-SummaryLine 'LDAP pages retrieved' (Format-Count $ldapMetrics.Pages)
            Write-SummaryLine 'LDAP entries returned' (Format-Count $ldapMetrics.Entries)
            Write-SummaryLine 'Duplicate AD SIDs' (Format-Count $ldapMetrics.DuplicateSids)
            Write-SummaryLine 'AD entries without SID' (Format-Count $ldapMetrics.MissingSids)
        }
        Write-SummaryLine 'Entra users with SID' (Format-Count $sidEligible.Count)
        Write-SummaryLine 'Entra users without SID' (Format-Count $noSid)
        Write-SummaryLine 'Entra users matched to AD' (Format-Count $sidMatched)
        Write-SummaryLine 'Entra users unmatched' (Format-Count $sidUnmatched)
        Write-SummaryLine 'AD SID match coverage' $coveragePercent
        if ($ldapAvailable -and $sidUnmatched -gt 0) { Write-Warning 'Some Entra users did not match an AD user within the selected OU. Their password status remains Unknown.' }
        Write-Host "`nAD Password Expiration (processed enabled users)" -ForegroundColor Cyan
        Write-SummaryLine 'AD users indexed' (Format-Count $adPasswords.Count)
        Write-SummaryLine 'Enabled passwords expired' (Format-Count @($results | Where-Object { $_.AccountEnabled -eq $true -and $_.PasswordStatus -in @('Expired','Must Change Password') }).Count)
        Write-SummaryLine 'Enabled passwords unknown' (Format-Count @($results | Where-Object { $_.AccountEnabled -eq $true -and $_.PasswordStatus -in @('Unknown','Invalid') }).Count)
        Write-SummaryLine 'Known password statuses' (Format-Count @($results | Where-Object { $_.PasswordStatus -in @('Active','Expired','Never','Must Change Password') }).Count)
        Write-Host '7/14/30-day store columns are cumulative and exclude already expired passwords.' -ForegroundColor DarkGray
    }

    Write-Host "`nPerformance`n=========== " -ForegroundColor Cyan
    foreach ($name in $timings.Keys) { Write-Host ('{0,-29}: {1,12}' -f $name, $timings[$name].ToString('hh\:mm\:ss\.fff')) }
    Write-Host ('{0,-29}: {1,12}' -f 'Total elapsed', $overall.Elapsed.ToString('hh\:mm\:ss\.fff'))
    Write-Host "`nUser report: $outputCsv" -ForegroundColor Cyan
    Write-Host "Store report: $storeCsv" -ForegroundColor Cyan
} catch {
    Write-Error ('Report failed: {0}' -f $_.Exception.Message)
    throw
}
