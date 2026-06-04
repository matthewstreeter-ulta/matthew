# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All"

# Find all App Registrations
[System.Collections.Generic.List[PSObject]]$appSecretsArray = @()

# We explicitly request PasswordCredentials as it is not always returned by default in large batches
$apps = Get-MgApplication -All -Property "DisplayName", "AppId", "Id", "PasswordCredentials"

foreach ($app in $apps) {
    # An application can have multiple client secrets
    foreach ($secret in $app.PasswordCredentials) {
        $object = [PSCustomObject][ordered]@{
            DisplayName        = $app.DisplayName
            AppId              = $app.AppId
            ObjectId           = $app.Id
            SecretDisplayName  = $secret.DisplayName
            SecretId           = $secret.KeyId
            CreatedDateTime    = $secret.StartDateTime
            ExpiryDateTime     = $secret.EndDateTime
            # Calculate if the secret is valid (not expired)
            IsExpired          = $secret.EndDateTime -lt (Get-Date)
            # Calculate remaining days
            DaysUntilExpiry    = (New-TimeSpan -Start (Get-Date) -End $secret.EndDateTime).Days
        }

        $appSecretsArray.Add($object)
    }
}

# Export to CSV
$dateTime = Get-Date -Format "yyyMMdd-HHmmss"
$outputPath = ".\Temp\AppSecrets-$dateTime.csv"
if (-not (Test-Path ".\Temp")) { New-Item -Path ".\Temp" -ItemType Directory | Out-Null }
$appSecretsArray | Export-Csv -Path $outputPath -Encoding UTF8 -NoClobber -NoTypeInformation 