# ============================================================
# Security Question Registration Report
#
# Uses an exported Entra security group membership CSV to
# determine which users have Security Questions registered
# as an authentication method.
#
# Required CSV columns:
#   Id
#   userPrincipalName
#   displayName
#   mail
# ============================================================


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

$InputCsv = "\Path\To\Input\Members.csv"

$OutputCsv = "\Path\To\Output\Members-SecurityQuestionRegistration.csv"

# Number of registration records requested from Graph per page
$pageSize = 1000


# ------------------------------------------------------------
# Start overall timer
# ------------------------------------------------------------

$scriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()


# ------------------------------------------------------------
# Import group membership
# ------------------------------------------------------------

Write-Host ""
Write-Host "Importing group members..." -ForegroundColor Cyan

$importStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$groupMembers = Import-Csv -Path $InputCsv

$importStopwatch.Stop()

Write-Host (
    "Group members imported: {0:N0}" -f
    $groupMembers.Count
) -ForegroundColor Cyan


# ------------------------------------------------------------
# Validate CSV
# ------------------------------------------------------------

if ($groupMembers.Count -eq 0) {
    throw "The input CSV contains no users."
}

$csvColumns = $groupMembers[0].PSObject.Properties.Name

$requiredColumns = @(
    'Id'
    'userPrincipalName'
    'displayName'
    'mail'
)

foreach ($column in $requiredColumns) {

    if ($column -notin $csvColumns) {
        throw "Required column '$column' was not found in the input CSV."
    }
}

Write-Host "CSV columns validated." -ForegroundColor Green


# ------------------------------------------------------------
# Retrieve authentication registration report
# ------------------------------------------------------------

Write-Host ""
Write-Host "Retrieving authentication method registration report..." `
    -ForegroundColor Cyan

Write-Host ""

$graphStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Generic list is more efficient than repeatedly using +=
$registrationDetails =
    [System.Collections.Generic.List[object]]::new()

$pageNumber = 0

# Graph endpoint for authentication method registration details
$uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?`$top=$pageSize"


# ------------------------------------------------------------
# Retrieve Graph report one page at a time
# ------------------------------------------------------------

while (-not [string]::IsNullOrWhiteSpace($uri)) {

    $pageNumber++

    # Retrieve one page
    $response = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $uri `
        -OutputType PSObject

    # Add returned records to our collection
    foreach ($record in $response.value) {
        $registrationDetails.Add($record)
    }

    $recordsRetrieved = $registrationDetails.Count
    $elapsed = $graphStopwatch.Elapsed


    # --------------------------------------------------------
    # Calculate average Graph retrieval rate
    # --------------------------------------------------------

    if ($elapsed.TotalSeconds -gt 0) {

        $recordsPerSecond = [math]::Round(
            $recordsRetrieved / $elapsed.TotalSeconds,
            0
        )
    }
    else {
        $recordsPerSecond = 0
    }


    # --------------------------------------------------------
    # Display retrieval progress
    # --------------------------------------------------------

    $progressMessage = (
        "Retrieved {0:N0} records | Page {1:N0} | " +
        "{2:N0} records/sec | Elapsed {3}"
    ) -f `
        $recordsRetrieved,
        $pageNumber,
        $recordsPerSecond,
        $elapsed.ToString("hh\:mm\:ss")

    Write-Host $progressMessage -ForegroundColor Yellow


    # --------------------------------------------------------
    # Get next page
    #
    # Graph supplies @odata.nextLink when additional records
    # remain. If it is absent, the loop ends.
    # --------------------------------------------------------

    $uri = $response.'@odata.nextLink'
}

$graphStopwatch.Stop()


# ------------------------------------------------------------
# Graph retrieval complete
# ------------------------------------------------------------

Write-Host ""
Write-Host "Registration retrieval complete." -ForegroundColor Green

Write-Host (
    "Records retrieved : {0:N0}" -f
    $registrationDetails.Count
)

Write-Host (
    "Elapsed           : {0}" -f
    $graphStopwatch.Elapsed.ToString("hh\:mm\:ss\.fff")
)


# ------------------------------------------------------------
# Build registration lookup table
# ------------------------------------------------------------

Write-Host ""
Write-Host "Building registration lookup table..." `
    -ForegroundColor Cyan

$lookupStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$registrationById = @{}

foreach ($reg in $registrationDetails) {

    if (-not [string]::IsNullOrWhiteSpace($reg.Id)) {

        $registrationById[$reg.Id] = $reg
    }
}

$lookupStopwatch.Stop()

Write-Host (
    "Lookup table complete: {0:N0} users" -f
    $registrationById.Count
) -ForegroundColor Green


# ------------------------------------------------------------
# Process group members
# ------------------------------------------------------------

Write-Host ""
Write-Host "Checking group members..." -ForegroundColor Cyan

$processingStopwatch =
    [System.Diagnostics.Stopwatch]::StartNew()


# ------------------------------------------------------------
# Counters
# ------------------------------------------------------------

$totalUsers      = $groupMembers.Count
$processedUsers  = 0
$matchedUsers    = 0
$unmatchedUsers  = 0
$securityUsers   = 0
$noSecurityUsers = 0


# ------------------------------------------------------------
# Compare group members against registration report
# ------------------------------------------------------------

$results = foreach ($member in $groupMembers) {

    $processedUsers++

    # Values supplied by the Entra group membership CSV
    $objectId    = $member.Id
    $upn         = $member.userPrincipalName
    $displayName = $member.displayName
    $mail        = $member.mail

    $regDetails = $null


    # --------------------------------------------------------
    # Find registration record using Entra Object ID
    # --------------------------------------------------------

    if (
        -not [string]::IsNullOrWhiteSpace($objectId) -and
        $registrationById.ContainsKey($objectId)
    ) {

        $regDetails = $registrationById[$objectId]

        $matchedUsers++
    }
    else {

        $unmatchedUsers++
    }


    # --------------------------------------------------------
    # Check Security Question registration
    # --------------------------------------------------------

    $hasSecurityQuestions = $false

    if (
        $null -ne $regDetails -and
        $regDetails.MethodsRegistered -contains 'securityQuestion'
    ) {

        $hasSecurityQuestions = $true

        $securityUsers++
    }
    elseif ($null -ne $regDetails) {

        # Only count this as "No Security Questions" when a
        # registration record was actually found.
        $noSecurityUsers++
    }


    # --------------------------------------------------------
    # Output result object
    # --------------------------------------------------------

    [PSCustomObject]@{

        ObjectId =
            $objectId

        DisplayName =
            $displayName

        UserPrincipalName =
            $upn

        Mail =
            $mail

        SecurityQuestionsRegistered =
            $hasSecurityQuestions

        MethodsRegistered =
            if ($null -ne $regDetails) {

                $regDetails.MethodsRegistered -join '; '
            }
            else {

                $null
            }

        RegistrationRecordFound =
            ($null -ne $regDetails)
    }
}

$processingStopwatch.Stop()


# ------------------------------------------------------------
# Group comparison complete
# ------------------------------------------------------------

Write-Host (
    "Group comparison complete: {0:N0} users checked in {1}" -f
    $totalUsers,
    $processingStopwatch.Elapsed.ToString("hh\:mm\:ss\.fff")
) -ForegroundColor Green


# ------------------------------------------------------------
# Export report
# ------------------------------------------------------------

Write-Host ""
Write-Host "Exporting results..." -ForegroundColor Cyan

$exportStopwatch =
    [System.Diagnostics.Stopwatch]::StartNew()

$results |
    Export-Csv `
        -Path $OutputCsv `
        -NoTypeInformation `
        -Encoding UTF8

$exportStopwatch.Stop()

$scriptStopwatch.Stop()


# ------------------------------------------------------------
# Final summary
# ------------------------------------------------------------

Write-Host ""
Write-Host "Report Complete" -ForegroundColor Green
Write-Host "===============" -ForegroundColor Green

Write-Host ""

Write-Host (
    "Group members             : {0:N0}" -f
    $totalUsers
)

Write-Host (
    "Registration records      : {0:N0}" -f
    $registrationDetails.Count
)

Write-Host (
    "Users matched             : {0:N0}" -f
    $matchedUsers
)

Write-Host (
    "Users not found           : {0:N0}" -f
    $unmatchedUsers
)

Write-Host (
    "Security Questions        : {0:N0}" -f
    $securityUsers
)

Write-Host (
    "No Security Questions     : {0:N0}" -f
    $noSecurityUsers
)


# ------------------------------------------------------------
# Performance summary
# ------------------------------------------------------------

Write-Host ""
Write-Host "Performance" -ForegroundColor Cyan
Write-Host "===========" -ForegroundColor Cyan

Write-Host ""

Write-Host (
    "CSV import                : {0:hh\:mm\:ss\.fff}" -f
    $importStopwatch.Elapsed
)

Write-Host (
    "Graph retrieval           : {0:hh\:mm\:ss\.fff}" -f
    $graphStopwatch.Elapsed
)

Write-Host (
    "Lookup table build        : {0:hh\:mm\:ss\.fff}" -f
    $lookupStopwatch.Elapsed
)

Write-Host (
    "Group comparison          : {0:hh\:mm\:ss\.fff}" -f
    $processingStopwatch.Elapsed
)

Write-Host (
    "CSV export                : {0:hh\:mm\:ss\.fff}" -f
    $exportStopwatch.Elapsed
)

Write-Host (
    "Total elapsed             : {0:hh\:mm\:ss\.fff}" -f
    $scriptStopwatch.Elapsed
)


# ------------------------------------------------------------
# Output location
# ------------------------------------------------------------

Write-Host ""
Write-Host "Output:" -ForegroundColor Cyan
Write-Host $OutputCsv
Write-Host ""
