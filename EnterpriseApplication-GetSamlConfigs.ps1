# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All"

# Find all SAML Applications
[System.Collections.Generic.List[PSObject]]$samlApplicationsArray = @()
$samlApplications = Get-MgBetaServicePrincipal -Filter "PreferredSingleSignOnMode eq 'saml'" -All # PreferredSingleSignOnMode = saml, oidc, password, OR notSupported

foreach ($samlApp in $samlApplications) {
    $object = [PSCustomObject][ordered]@{
        DisplayName                         = $samlApp.DisplayName
        Id                                  = $samlApp.Id
        AppId                               = $samlApp.AppId
        LoginUrl                            = $samlApp.LoginUrl
        LogoutUrl                           = $samlApp.LogoutUrl
        NotificationEmailAddresses          = $samlApp.NotificationEmailAddresses -join '|'
        AppRoleAssignmentRequired           = $samlApp.AppRoleAssignmentRequired
        PreferredSingleSignOnMode           = $samlApp.PreferredSingleSignOnMode
        PreferredTokenSigningKeyEndDateTime = $samlApp.PreferredTokenSigningKeyEndDateTime
        # PreferredTokenSigningKeyEndDateTime is date time, compared to now and see it is valid
        PreferredTokenSigningKeyValid       = $samlApp.PreferredTokenSigningKeyEndDateTime -gt (Get-Date)
        ReplyUrls                           = $samlApp.ReplyUrls -join '|'
        SignInAudience                      = $samlApp.SignInAudience
    }

    $samlApplicationsArray.Add($object)
}

#return $samlApplicationsArray

# Export to CSV
$dateTime = Get-Date -Format "yyyMMdd-HHmmss"
$samlApplicationsArray | Export-Csv -Path ".\Temp\SamlApps-$dateTime.csv" -Encoding UTF8 -NoClobber -NoTypeInformation 