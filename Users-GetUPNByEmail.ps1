# Define file paths
$InputCsv = "\path\to\input.csv"
$OutputCsv = "\path\to\output.csv"

# Connect to Microsoft Graph
# This will prompt you to sign in. You need an account with permissions to read users.
#Connect-MgGraph -Scopes "User.Read.All"

# Check for CSV file
if (-not (Test-Path $InputCsv)) {
    Write-Error "Cannot find the input CSV file at $InputCsv"
    exit
}

# Import CSV file
$CsvData = Import-Csv -Path $InputCsv

# Process each row and query Graph
$Results = foreach ($Row in $CsvData) {
    $Email = $Row.Email.Trim()
    
    if (-not [string]::IsNullOrWhiteSpace($Email)) {
        Write-Host "Auditing: $Email" -ForegroundColor Cyan

        # Fetch extra account properties to differentiate primary vs admin/secondary accounts
        $GraphUsers = Get-MgUser -Filter "mail eq '$Email' or userPrincipalName eq '$Email'" `
            -Property "id, displayName, userPrincipalName, mail, userType, accountEnabled, createdDateTime" `
            -ErrorAction SilentlyContinue

        if ($GraphUsers) {
            # Loop through EACH matching user so every account gets its own row in the report
            foreach ($User in $GraphUsers) {
                [PSCustomObject]@{
                    SearchEmail       = $Email
                    TotalMatchesFound = $GraphUsers.Count
                    IsDuplicate       = if ($GraphUsers.Count -gt 1) { $true } else { $false }
                    DisplayName       = $User.DisplayName
                    UserPrincipalName = $User.UserPrincipalName
                    Mail              = $User.Mail
                    AccountEnabled    = $User.AccountEnabled
                    UserType          = $User.UserType
                    CreatedDateTime   = $User.CreatedDateTime
                    Status            = if ($GraphUsers.Count -gt 1) { "Duplicate - Action Needed" } else { "Unique Match" }
                }
            }
        } else {
            [PSCustomObject]@{
                SearchEmail       = $Email
                TotalMatchesFound = 0
                IsDuplicate       = $false
                DisplayName       = $null
                UserPrincipalName = $null
                Mail              = $null
                AccountEnabled    = $null
                UserType          = $null
                CreatedDateTime   = $null
                Status            = "Not Found"
            }
        }
    }
}

# Export full audit list
$Results | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Audit completed! Report written to $OutputCsv" -ForegroundColor Green

# Clean up the session
#Disconnect-MgGraph