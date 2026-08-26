# Requires an input users csv file with column headers: "EE ID", "First Name", "Last Name"
# First and Last are fallback options if the user has an older SAM/UPN standard (non-numbers)

# Connect to Microsoft Graph. 
# Note: The reporting cmdlet requires the 'AuditLog.Read.All' scope in addition to 'User.Read.All'.
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"

# Define file paths
$csvInputPath = "\Path\To\Input\UsersList.csv"
$csvOutputPath = "\Path\To\Output\UsersList-SecurityQuestionsReport.csv"
$summaryOutputPath = "\Path\To\Output\UsersList-SecurityQuestionsSummary.txt"

# Read the CSV file containing the 'EE ID', 'First Name', and 'Last Name' columns
if (-not (Test-Path $csvInputPath)) {
    Write-Error "Input CSV file not found at: $csvInputPath"
    exit
}
$userList = Import-Csv -Path $csvInputPath
$totalUsers = $userList.Count
$currentIndex = 1

$report = @()

Write-Host "Starting Security Questions check for $totalUsers users..." -ForegroundColor Cyan

foreach ($row in $userList) {
    # Dynamically locate the 'EE ID' property key to bypass hidden BOM/UTF-8 control characters
    $eeIdKey   = $row.PSObject.Properties.Name | Where-Object { $_ -like "*EE ID*" } | Select-Object -First 1
    $empId     = $row.$eeIdKey
    $firstName = $row.'First Name'
    $lastName  = $row.'Last Name'
    
    if ([string]::IsNullOrWhiteSpace($empId) -and [string]::IsNullOrWhiteSpace($firstName)) {
        $currentIndex++
        continue
    }
    
    $user = $null
    $matchMethod = "None"

    # PRIMARY SEARCH: Try matching by onPremisesSamAccountName (EE ID)
    if (-not [string]::IsNullOrWhiteSpace($empId)) {
        $user = Get-MgUser -Filter "onPremisesSamAccountName eq '$empId'" -ConsistencyLevel eventual -CountVariable userCount -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($user) { $matchMethod = "EE ID" }
    }

    # FALLBACK SEARCH: Try matching by First Name and Last Name if ID fails
    if (-not $user -and -not [string]::IsNullOrWhiteSpace($firstName) -and -not [string]::IsNullOrWhiteSpace($lastName)) {
        # Escape single quotes in names (e.g., O'Brien -> O''Brien) to prevent Graph API syntax errors
        $safeFirst = $firstName -replace "'", "''"
        $safeLast  = $lastName -replace "'", "''"
        
        $user = Get-MgUser -Filter "givenName eq '$safeFirst' and surname eq '$safeLast'" -ConsistencyLevel eventual -CountVariable userCount -Property Id, UserPrincipalName, DisplayName -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($user) { $matchMethod = "Name Match" }
    }

    if ($user) {
        $upn = $user.UserPrincipalName
        $displayName = $user.DisplayName
        
        if ($matchMethod -eq "Name Match") {
            Write-Host "[$currentIndex/$totalUsers] Processing: $displayName " -NoNewline
            Write-Host "(Name Fallback)" -ForegroundColor Cyan -NoNewline
        } else {
            Write-Host "[$currentIndex/$totalUsers] Processing: $displayName ($empId)" -NoNewline
        }

        # Query the authentication method registration details report for this specific UPN
        $regDetails = Get-MgReportAuthenticationMethodUserRegistrationDetail -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
        
        # Check if the 'MethodsRegistered' array contains 'securityQuestion'
        $hasSecurityQuestions = $false
        if ($null -ne $regDetails -and $regDetails.MethodsRegistered -contains 'securityQuestion') {
            $hasSecurityQuestions = $true
        }

        if ($hasSecurityQuestions) {
            Write-Host " -> Security Questions registered." -ForegroundColor Green
        } else {
            Write-Host " -> NO Security Questions registered." -ForegroundColor Red
        }

        $report += [PSCustomObject]@{
            'EE ID'                = $empId
            'MatchMethod'          = $matchMethod
            'DisplayName'          = $displayName
            'UserPrincipalName'    = $upn
            'HasSecurityQuestions' = $hasSecurityQuestions
        }
    }
    else {
        # Handle cases where the user does not exist in Entra ID by either ID or Name
        Write-Host "[$currentIndex/$totalUsers] Processing: EE ID $empId" -NoNewline
        Write-Host " -> NOT FOUND in Entra ID." -ForegroundColor Yellow

        $report += [PSCustomObject]@{
            'EE ID'                = $empId
            'MatchMethod'          = "Not Found"
            'DisplayName'          = "$firstName $lastName"
            'UserPrincipalName'    = "Not Found in Entra"
            'HasSecurityQuestions' = "N/A"
        }
    }
    
    $currentIndex++
}

# Export the results array to a new CSV
Write-Host "`nProcessing complete. Exporting results for $($report.Count) users..." -ForegroundColor Cyan
$report | Export-Csv -Path $csvOutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Export completed successfully to: $csvOutputPath" -ForegroundColor Green

# ---------------------------------------------------------
# GENERATE SUMMARY
# ---------------------------------------------------------
$trueCount      = @($report | Where-Object { $_.HasSecurityQuestions -eq $true }).Count
$falseCount     = @($report | Where-Object { $_.HasSecurityQuestions -eq $false }).Count
$notFoundCount  = @($report | Where-Object { $_.MatchMethod -eq 'Not Found' }).Count
$nameMatchCount = @($report | Where-Object { $_.MatchMethod -eq 'Name Match' }).Count

$percentTrue = 0
if ($totalUsers -gt 0) {
    $percentTrue = [math]::Round((($trueCount / $totalUsers) * 100), 2)
}

# Console Output
Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "              REPORT SUMMARY             " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Total Users Processed : $totalUsers"
Write-Host "Registered (True)     : $trueCount ($percentTrue%)" -ForegroundColor Green
Write-Host "Not Registered (False): $falseCount" -ForegroundColor Red
if ($notFoundCount -gt 0) {
    Write-Host "Not Found in Entra ID : $notFoundCount" -ForegroundColor Yellow
}
if ($nameMatchCount -gt 0) {
    Write-Host "Matched by Name (FB)  : $nameMatchCount" -ForegroundColor Cyan
}
Write-Host "=========================================`n" -ForegroundColor Cyan

# Text File Output
$summaryText = @"
=========================================
              REPORT SUMMARY             
=========================================
Total Users Processed : $totalUsers
Registered (True)     : $trueCount ($percentTrue%)
Not Registered (False): $falseCount
"@

if ($notFoundCount -gt 0) {
    $summaryText += "`r`nNot Found in Entra ID : $notFoundCount"
}
if ($nameMatchCount -gt 0) {
    $summaryText += "`r`nMatched by Name (FB)  : $nameMatchCount"
}
$summaryText += "`r`n========================================="

$summaryText | Out-File -FilePath $summaryOutputPath -Encoding UTF8
Write-Host "Summary statistics saved to: $summaryOutputPath" -ForegroundColor Green
