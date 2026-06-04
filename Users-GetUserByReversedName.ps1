# This script takes a list of names in "First Last" format, 
# converts them to "Last, First" if the Entra tenant is configured
# with DisplayName in that format , and retrieves their UPN and Email from Entra.

# Set to $true if you prefer the pop-up window, $false to stay in the command line
$UseGridView = $false

# Path to your input list
$InputFilePath = "C:\Users\9118123\OneDrive - Ulta\Scripts\Temp\Names.txt"
$OutputCsvPath = "C:\Users\9118123\OneDrive - Ulta\Scripts\Temp\UserLookupResults-$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"

# Ensure Microsoft Graph connection
# Connect-MgGraph -Scopes "User.Read.All"

if (-not (Test-Path $InputFilePath)) {
    Write-Error "Input file not found at $InputFilePath"
    return
}

$names = Get-Content -Path $InputFilePath
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Starting lookup for $($names.Count) users..." -ForegroundColor Cyan

foreach ($rawName in $names) {
    $trimmedName = $rawName.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedName)) { continue }

    # Split "First Last" into parts
    # This regex splits by the last space to handle middle names (e.g., "Mary Jo Smith" -> "Smith, Mary Jo")
    if ($trimmedName -match '^(.*)\s+(.*)$') {
        $firstName = $Matches[1]
        $lastName  = $Matches[2]
        $searchDisplayName = "$lastName, $firstName"
    } else {
        Write-Warning "Could not parse name format for: $trimmedName. Skipping."
        continue
    }

    Write-Host "Searching for: $searchDisplayName" -NoNewline

    try {
        $user = $null
        $status = "Not Found"

        # Filter using the "Last, First" format
        $user = Get-MgUser -Filter "displayName eq '$searchDisplayName'" -Property UserPrincipalName, Mail, DisplayName, Id -ErrorAction Stop
        
        if ($user) {
            $status = "Found (Exact)"
        } else {
            # Step 2: Candidate search by Last Name
            Write-Host " -> No exact match. Searching candidates for '$lastName'..." -ForegroundColor Yellow
            # Force result to array with @() and filter out null DisplayNames
            $candidates = @(Get-MgUser -Filter "startsWith(displayName, '$lastName') or surname eq '$lastName'" -Property UserPrincipalName, Mail, DisplayName, Id -Top 15 | Where-Object { $_.DisplayName })
            
            if ($candidates.Count -gt 0) {
                if ($UseGridView) {
                    $user = $candidates | Out-GridView -Title "Select correct user for '$trimmedName'" -PassThru
                } else {
                    Write-Host "`nMultiple candidates found for '$trimmedName':" -ForegroundColor Cyan
                    for ($i = 0; $i -lt $candidates.Count; $i++) {
                        Write-Host "  [$($i + 1)] $($candidates[$i].DisplayName) ($($candidates[$i].UserPrincipalName))"
                    }
                    $choice = Read-Host "`nSelect a number (1-$($candidates.Count)) or press Enter to skip"
                    if ($choice -as [int] -and $choice -ge 1 -and $choice -le $candidates.Count) {
                        $user = $candidates[$choice - 1]
                    }
                }
                if ($user) { $status = "Found (Candidate Selection)" }
            }

            # Step 3: Manual search fallback
            if (-not $user) {
                $manualTerm = Read-Host " -> No candidate selected. Enter partial name to search manually (or Enter to skip)"
                if (-not [string]::IsNullOrWhiteSpace($manualTerm)) {
                    $manualCandidates = @(Get-MgUser -Filter "startsWith(displayName, '$manualTerm') or startsWith(givenName, '$manualTerm') or startsWith(surname, '$manualTerm')" -Property UserPrincipalName, Mail, DisplayName, Id -Top 15 | Where-Object { $_.DisplayName })
                    if ($manualCandidates.Count -gt 0) {
                        Write-Host "`nResults for manual search '$manualTerm':" -ForegroundColor Cyan
                        for ($i = 0; $i -lt $manualCandidates.Count; $i++) {
                            Write-Host "  [$($i + 1)] $($manualCandidates[$i].DisplayName) ($($manualCandidates[$i].UserPrincipalName))"
                        }
                        $choice = Read-Host "`nSelect a number (1-$($manualCandidates.Count)) or press Enter to skip"
                        if ($choice -as [int] -and $choice -ge 1 -and $choice -le $manualCandidates.Count) {
                            $user = $manualCandidates[$choice - 1]
                        }
                    }
                    if ($user) { $status = "Found (Manual Search)" }
                }
            }
        }

        if ($user) {
            Write-Host " -> Found: $($user.DisplayName)" -ForegroundColor Green
            $results.Add([PSCustomObject]@{
                InputName         = $trimmedName
                MatchedName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Email             = $user.Mail
                Status            = $status
            })
        } else {
            $results.Add([PSCustomObject]@{ InputName = $trimmedName; Status = "Not Found" })
        }
    } catch {
        Write-Warning " -> Error: $($_.Exception.Message)"
    }
}

$results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "`nProcess complete. Results saved to: $OutputCsvPath" -ForegroundColor Cyan
