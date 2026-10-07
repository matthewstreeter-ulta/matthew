<#
.SYNOPSIS
    Remote Entra Connect Sync utility for authorized IAM administrators.

.DESCRIPTION
    Provides a menu for:
      - Starting an Entra Connect Delta synchronization remotely
      - Monitoring live synchronization activity
      - Checking synchronization and scheduler status
      - Creating/updating a locally stored privileged credential
      - Testing remote connectivity

    Credentials are stored using Export-Clixml.

    On Windows, the password contained in the exported PSCredential is
    protected using Windows DPAPI. It can only be decrypted by the same
    Windows user on the same computer that created it.

    This script does NOT modify the Entra Connect scheduler.
#>


# ============================================================
# CONFIGURATION
# ============================================================

$SyncServer = "ENTRA-CONNECT-SERVERNAME"

$CredentialDirectory = Join-Path $env:APPDATA "IAM"
$CredentialPath = Join-Path $CredentialDirectory "EntraConnectSync.xml"

# Live sync monitoring settings
$SyncPollIntervalSeconds = 1
$SyncStartupTimeoutSeconds = 60
$SyncCompletionTimeoutMinutes = 15


# ============================================================
# DISPLAY FUNCTIONS
# ============================================================

function Write-Header {

    Clear-Host

    Write-Host "============================================================"
    Write-Host " Entra Connect Remote Sync Utility"
    Write-Host "============================================================"
    Write-Host ""
    Write-Host "Sync Server : $SyncServer"
    Write-Host "Workstation : $env:COMPUTERNAME"
    Write-Host "User        : $env:USERDOMAIN\$env:USERNAME"
    Write-Host ""
}


function Wait-ForEnter {

    Write-Host ""
    Read-Host "Press Enter to continue"
}


# ============================================================
# CREDENTIAL FUNCTIONS
# ============================================================

function Update-EntraSyncCredential {

    Write-Host ""
    Write-Host "Stored Credential Setup"
    Write-Host "-----------------------"
    Write-Host ""
    Write-Host "Enter the privileged account used to remotely administer"
    Write-Host "the Entra Connect server."
    Write-Host ""

    $Cred = Get-Credential -Message "Enter your Entra Connect administrative credential"

    if ($null -eq $Cred) {

        Write-Host ""
        Write-Warning "Credential update cancelled."

        return $false
    }

    try {

        if (-not (Test-Path -Path $CredentialDirectory)) {
            New-Item -ItemType Directory -Path $CredentialDirectory -Force | Out-Null
        }

        $Cred | Export-Clixml -Path $CredentialPath -Force

        Write-Host ""
        Write-Host "Stored credential updated successfully."
        Write-Host ""
        Write-Host "Account         : $($Cred.UserName)"
        Write-Host "Credential File : $CredentialPath"

        return $true
    }
    catch {

        Write-Host ""
        Write-Error "Unable to save the credential."
        Write-Host $_.Exception.Message

        return $false
    }
}


function Get-EntraSyncCredential {

    if (-not (Test-Path -Path $CredentialPath)) {

        Write-Host ""
        Write-Warning "No stored Entra Connect credential was found."
        Write-Host ""
        Write-Host "A credential must be configured before remote commands can run."
        Write-Host ""

        $Response = Read-Host "Create the stored credential now? [Y/N]"

        if ($Response -match "^[Yy]$") {

            $Created = Update-EntraSyncCredential

            if (-not $Created) {
                return $null
            }
        }
        else {
            return $null
        }
    }

    try {

        $Cred = Import-Clixml -Path $CredentialPath

        return $Cred
    }
    catch {

        Write-Host ""
        Write-Error "Unable to read the stored credential."
        Write-Host ""
        Write-Host "The credential may have been created by another Windows"
        Write-Host "account or on another computer."
        Write-Host ""

        return $null
    }
}


# ============================================================
# CONNECTION FUNCTIONS
# ============================================================

function Test-EntraSyncConnection {

    param (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host ""
    Write-Host "Testing connection to $SyncServer..."

    $InvokeParams = @{
        ComputerName = $SyncServer
        Credential   = $Credential
        ScriptBlock  = {

            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                User         = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            }
        }
        ErrorAction = "Stop"
    }

    try {

        $Result = Invoke-Command @InvokeParams

        Write-Host "Connection successful."
        Write-Host ""

        return $Result
    }
    catch {

        Write-Host ""
        Write-Warning "Unable to authenticate or connect to $SyncServer."
        Write-Host ""
        Write-Host "Possible causes:"
        Write-Host "  - Your privileged account password has changed"
        Write-Host "  - The account is locked or disabled"
        Write-Host "  - WinRM is unavailable"
        Write-Host "  - Remote PowerShell access has changed"
        Write-Host ""
        Write-Host "Error:"
        Write-Host $_.Exception.Message
        Write-Host ""

        return $null
    }
}


function Invoke-CredentialRecovery {

    Write-Host "Would you like to update your stored credential?"
    Write-Host ""

    $Response = Read-Host "Update credential? [Y/N]"

    if ($Response -match "^[Yy]$") {

        $Updated = Update-EntraSyncCredential

        if ($Updated) {

            Write-Host ""
            Write-Host "Testing the updated credential..."

            $NewCred = Get-EntraSyncCredential

            if ($null -ne $NewCred) {

                $Connection = Test-EntraSyncConnection -Credential $NewCred

                if ($null -ne $Connection) {

                    Write-Host "Updated credential successfully authenticated."
                    Write-Host ""

                    return $NewCred
                }
            }
        }
    }

    return $null
}


function Get-ValidatedCredential {

    $Cred = Get-EntraSyncCredential

    if ($null -eq $Cred) {
        return $null
    }

    $Connection = Test-EntraSyncConnection -Credential $Cred

    if ($null -ne $Connection) {
        return $Cred
    }

    $Cred = Invoke-CredentialRecovery

    return $Cred
}


# ============================================================
# LIVE SYNC ENGINE STATE
# ============================================================

function Get-EntraSyncEngineState {

    param (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )

    $InvokeParams = @{
        ComputerName = $SyncServer
        Credential   = $Credential
        ScriptBlock  = {

            Import-Module ADSync

            $RunStatus = Get-ADSyncConnectorRunStatus

            if ($RunStatus) {

                [PSCustomObject]@{
                    IsRunning     = $true
                    RunState      = $RunStatus.RunState.ToString()
                    ConnectorName = $RunStatus.ConnectorName
                }
            }
            else {

                [PSCustomObject]@{
                    IsRunning     = $false
                    RunState      = "Idle"
                    ConnectorName = $null
                }
            }
        }
        ErrorAction = "Stop"
    }

    Invoke-Command @InvokeParams
}


# ============================================================
# SYNC STATUS
# ============================================================

function Get-EntraSyncStatus {

    param (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host ""
    Write-Host "Retrieving Entra Connect synchronization status..."
    Write-Host ""

    $InvokeParams = @{
        ComputerName = $SyncServer
        Credential   = $Credential
        ScriptBlock  = {

            Import-Module ADSync

            $Scheduler = Get-ADSyncScheduler
            $RunStatus = Get-ADSyncConnectorRunStatus

            if ($RunStatus) {
                $EngineStatus = $RunStatus.RunState.ToString()
                $ConnectorName = $RunStatus.ConnectorName
            }
            else {
                $EngineStatus = "Idle"
                $ConnectorName = $null
            }

            [PSCustomObject]@{
                Server                = $env:COMPUTERNAME
                SyncEngineStatus      = $EngineStatus
                CurrentConnector      = $ConnectorName
                SchedulerEnabled      = $Scheduler.SyncCycleEnabled
                NextSyncCyclePolicy   = $Scheduler.NextSyncCyclePolicyType.ToString()
                NextSyncCycleStartUTC = $Scheduler.NextSyncCycleStartTimeInUTC
                EffectiveSyncInterval = $Scheduler.CurrentlyEffectiveSyncCycleInterval
            }
        }
        ErrorAction = "Stop"
    }

    try {

        $Status = Invoke-Command @InvokeParams

        $NextSyncUTC = $Status.NextSyncCycleStartUTC

        if ($null -ne $NextSyncUTC) {

            if ($NextSyncUTC.Kind -eq [System.DateTimeKind]::Unspecified) {

                $NextSyncUTC = [DateTime]::SpecifyKind(
                    $NextSyncUTC,
                    [System.DateTimeKind]::Utc
                )
            }

            $NextSyncLocal = $NextSyncUTC.ToLocalTime()
        }
        else {
            $NextSyncLocal = $null
        }

        Write-Host "Entra Connect Status"
        Write-Host "--------------------"
        Write-Host ""

        Write-Host "Server                  : $($Status.Server)"
        Write-Host "Sync Engine Status      : $($Status.SyncEngineStatus)"

        if ($Status.CurrentConnector) {
            Write-Host "Current Connector       : $($Status.CurrentConnector)"
        }

        Write-Host "Scheduler Enabled       : $($Status.SchedulerEnabled)"
        Write-Host "Next Sync Type          : $($Status.NextSyncCyclePolicy)"

        if ($null -ne $NextSyncLocal) {

            Write-Host "Next Scheduled Sync     : $($NextSyncLocal.ToString('MM/dd/yyyy hh:mm:ss tt'))"
            Write-Host "Next Scheduled Sync UTC : $($NextSyncUTC.ToString('MM/dd/yyyy HH:mm:ss'))"
        }
        else {

            Write-Host "Next Scheduled Sync     : Not available"
            Write-Host "Next Scheduled Sync UTC : Not available"
        }

        Write-Host "Effective Sync Interval : $($Status.EffectiveSyncInterval)"
        Write-Host ""
    }
    catch {

        Write-Host ""
        Write-Error "Unable to retrieve Entra Connect status."
        Write-Host $_.Exception.Message
        Write-Host ""
    }
}


# ============================================================
# LIVE DELTA SYNC MONITOR
# ============================================================

function Watch-EntraDeltaSync {

    param (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory)]
        [datetime]$RequestedAt
    )

    Write-Host ""
    Write-Host "Monitoring Delta Synchronization"
    Write-Host "================================"
    Write-Host ""

    $MonitoringStart = Get-Date
    $StartupDeadline = $MonitoringStart.AddSeconds($SyncStartupTimeoutSeconds)

    $SyncDetected = $false
    $LastConnector = $null

    $ActivityLog = [System.Collections.Generic.List[object]]::new()

    $ActivityLog.Add(
        [PSCustomObject]@{
            Time     = $RequestedAt
            Activity = "Delta synchronization request accepted"
        }
    )


    # ========================================================
    # PHASE 1 - WAIT FOR ENGINE TO BECOME ACTIVE
    # ========================================================

    while ((Get-Date) -lt $StartupDeadline) {

        try {

            $State = Get-EntraSyncEngineState -Credential $Credential
        }
        catch {

            Write-Progress -Activity "Entra Connect Delta Sync" -Completed

            Write-Host ""
            Write-Warning "Unable to query the synchronization engine while monitoring."
            Write-Host $_.Exception.Message
            Write-Host ""

            return
        }

        $Now = Get-Date
        $Elapsed = $Now - $RequestedAt

        if ($State.IsRunning) {

            $SyncDetected = $true

            $ActivityLog.Add(
                [PSCustomObject]@{
                    Time     = $Now
                    Activity = "Sync engine entered running state"
                }
            )

            if ($State.ConnectorName) {

                $LastConnector = $State.ConnectorName

                $ActivityLog.Add(
                    [PSCustomObject]@{
                        Time     = $Now
                        Activity = "Connector: $($State.ConnectorName)"
                    }
                )
            }

            Write-Progress -Activity "Entra Connect Delta Sync" -Completed

            Write-Host "Status            : RUNNING"
            Write-Host "Run State         : $($State.RunState)"

            if ($State.ConnectorName) {
                Write-Host "Current Connector : $($State.ConnectorName)"
            }

            Write-Host "Elapsed           : $($Elapsed.ToString('hh\:mm\:ss'))"
            Write-Host ""

            break
        }

        $ProgressParams = @{
            Activity = "Entra Connect Delta Sync"
            Status   = "Waiting for synchronization to start... Elapsed: $($Elapsed.ToString('hh\:mm\:ss'))"
        }

        Write-Progress @ProgressParams

        Start-Sleep -Seconds $SyncPollIntervalSeconds
    }


    # ========================================================
    # STARTUP TIMEOUT
    # ========================================================

    if (-not $SyncDetected) {

        Write-Progress -Activity "Entra Connect Delta Sync" -Completed

        Write-Host ""
        Write-Warning "The sync engine did not report a running connector within $SyncStartupTimeoutSeconds seconds."
        Write-Host ""
        Write-Host "The Delta request was accepted, but this monitoring session"
        Write-Host "could not confirm that the engine entered a running state."
        Write-Host ""
        Write-Host "Use 'Check Sync Status' to verify the current state."
        Write-Host ""

        return
    }


    # ========================================================
    # PHASE 2 - MONITOR ACTIVE SYNCHRONIZATION
    # ========================================================

    $CompletionDeadline = (Get-Date).AddMinutes($SyncCompletionTimeoutMinutes)

    while ((Get-Date) -lt $CompletionDeadline) {

        Start-Sleep -Seconds $SyncPollIntervalSeconds

        try {

            $State = Get-EntraSyncEngineState -Credential $Credential
        }
        catch {

            Write-Progress -Activity "Entra Connect Delta Sync" -Completed

            Write-Host ""
            Write-Warning "Unable to query the synchronization engine while monitoring."
            Write-Host $_.Exception.Message
            Write-Host ""

            return
        }

        $Now = Get-Date
        $Elapsed = $Now - $RequestedAt


        # ----------------------------------------------------
        # SYNC COMPLETED
        # ----------------------------------------------------

        if (-not $State.IsRunning) {

            Write-Progress -Activity "Entra Connect Delta Sync" -Completed

            $CompletedAt = Get-Date
            $TotalElapsed = $CompletedAt - $RequestedAt

            $ActivityLog.Add(
                [PSCustomObject]@{
                    Time     = $CompletedAt
                    Activity = "Sync engine returned to Idle"
                }
            )

            Write-Host ""
            Write-Host "Delta Synchronization Activity"
            Write-Host "=============================="
            Write-Host ""

            foreach ($Entry in $ActivityLog) {
                Write-Host "$($Entry.Time.ToString('hh:mm:ss tt'))  $($Entry.Activity)"
            }

            Write-Host ""
            Write-Host "Delta Synchronization Complete"
            Write-Host "=============================="
            Write-Host ""
            Write-Host "Server       : $SyncServer"
            Write-Host "Status       : Completed"
            Write-Host "Requested At : $($RequestedAt.ToString('MM/dd/yyyy hh:mm:ss tt'))"
            Write-Host "Completed At : $($CompletedAt.ToString('MM/dd/yyyy hh:mm:ss tt'))"
            Write-Host "Elapsed      : $($TotalElapsed.ToString('hh\:mm\:ss'))"
            Write-Host ""

            return
        }


        # ----------------------------------------------------
        # DETECT CONNECTOR TRANSITION
        # ----------------------------------------------------

        if (
            $State.ConnectorName -and
            $State.ConnectorName -ne $LastConnector
        ) {

            $LastConnector = $State.ConnectorName

            $ActivityLog.Add(
                [PSCustomObject]@{
                    Time     = $Now
                    Activity = "Connector: $($State.ConnectorName)"
                }
            )

            Write-Host "$($Now.ToString('hh:mm:ss tt'))  Connector changed: $($State.ConnectorName)"
        }


        # ----------------------------------------------------
        # LIVE STATUS
        # ----------------------------------------------------

        $StatusText = "Running"

        if ($State.ConnectorName) {
            $StatusText = "Running - $($State.ConnectorName)"
        }

        $ProgressParams = @{
            Activity = "Entra Connect Delta Sync"
            Status   = "$StatusText | Elapsed: $($Elapsed.ToString('hh\:mm\:ss'))"
        }

        Write-Progress @ProgressParams
    }


    # ========================================================
    # COMPLETION TIMEOUT
    # ========================================================

    Write-Progress -Activity "Entra Connect Delta Sync" -Completed

    $Elapsed = (Get-Date) - $RequestedAt

    Write-Host ""
    Write-Warning "The synchronization is still running after $SyncCompletionTimeoutMinutes minutes."
    Write-Host ""
    Write-Host "Monitoring has stopped."
    Write-Host ""
    Write-Host "This does NOT stop the Entra Connect synchronization."
    Write-Host "The synchronization engine will continue running normally."
    Write-Host ""
    Write-Host "Elapsed : $($Elapsed.ToString('hh\:mm\:ss'))"
    Write-Host ""
    Write-Host "Use 'Check Sync Status' to check it later."
    Write-Host ""
}


# ============================================================
# START DELTA SYNC
# ============================================================

function Start-EntraDeltaSync {

    param (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host ""
    Write-Host "Checking current synchronization state..."
    Write-Host ""

    $StatusParams = @{
        ComputerName = $SyncServer
        Credential   = $Credential
        ScriptBlock  = {

            Import-Module ADSync

            $Scheduler = Get-ADSyncScheduler
            $RunStatus = Get-ADSyncConnectorRunStatus

            [PSCustomObject]@{
                SyncRunning      = [bool]$RunStatus
                RunState         = if ($RunStatus) { $RunStatus.RunState.ToString() } else { "Idle" }
                CurrentConnector = if ($RunStatus) { $RunStatus.ConnectorName } else { $null }
                NextSyncStartUTC = $Scheduler.NextSyncCycleStartTimeInUTC
                NextSyncType     = $Scheduler.NextSyncCyclePolicyType.ToString()
            }
        }
        ErrorAction = "Stop"
    }

    try {

        $Status = Invoke-Command @StatusParams
    }
    catch {

        Write-Host ""
        Write-Error "Unable to query the synchronization engine."
        Write-Host $_.Exception.Message
        Write-Host ""

        return
    }


    # --------------------------------------------------------
    # PREVENT OVERLAPPING SYNC CYCLES
    # --------------------------------------------------------

    if ($Status.SyncRunning) {

        Write-Warning "An Entra Connect synchronization is already running."
        Write-Host ""
        Write-Host "Run State         : $($Status.RunState)"

        if ($Status.CurrentConnector) {
            Write-Host "Current Connector : $($Status.CurrentConnector)"
        }

        Write-Host ""
        Write-Host "No additional Delta synchronization was requested."
        Write-Host ""

        return
    }


    # --------------------------------------------------------
    # CONVERT SCHEDULER UTC TIME TO WORKSTATION LOCAL TIME
    # --------------------------------------------------------

    $NextSyncUTC = $Status.NextSyncStartUTC

    if ($null -ne $NextSyncUTC) {

        if ($NextSyncUTC.Kind -eq [System.DateTimeKind]::Unspecified) {

            $NextSyncUTC = [DateTime]::SpecifyKind(
                $NextSyncUTC,
                [System.DateTimeKind]::Utc
            )
        }

        $NextSyncLocal = $NextSyncUTC.ToLocalTime()
    }
    else {

        $NextSyncLocal = $null
    }


    # --------------------------------------------------------
    # DISPLAY PRE-SYNC STATUS
    # --------------------------------------------------------

    Write-Host "Sync Engine         : Idle"
    Write-Host "Next Scheduled Type : $($Status.NextSyncType)"

    if ($null -ne $NextSyncLocal) {
        Write-Host "Next Scheduled Sync : $($NextSyncLocal.ToString('MM/dd/yyyy hh:mm:ss tt'))"
    }
    else {
        Write-Host "Next Scheduled Sync : Not available"
    }

    Write-Host ""
    Write-Host "This will initiate an Entra Connect DELTA synchronization."
    Write-Host ""

    $Confirmation = Read-Host "Start Delta synchronization now? [Y/N]"

    if ($Confirmation -notmatch "^[Yy]$") {

        Write-Host ""
        Write-Host "Delta synchronization cancelled."
        Write-Host ""

        return
    }


    # --------------------------------------------------------
    # TRIGGER DELTA
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "Submitting Delta synchronization request..."
    Write-Host ""

    $SyncParams = @{
        ComputerName = $SyncServer
        Credential   = $Credential
        ScriptBlock  = {

            Import-Module ADSync

            $RequestedAt = Get-Date

            $SyncResult = Start-ADSyncSyncCycle -PolicyType Delta -InteractiveMode $false

            # Convert the ADSync-specific result to a normal string
            # BEFORE it crosses the PowerShell remoting boundary.
            $ResultString = $SyncResult.Result.ToString()

            [PSCustomObject]@{
                Server      = $env:COMPUTERNAME
                RequestedAt = $RequestedAt
                Result      = $ResultString
                WasAccepted = ($ResultString -eq "Success")
            }
        }
        ErrorAction = "Stop"
    }

    try {

        $Result = Invoke-Command @SyncParams

        Write-Host ""
        Write-Host "Delta synchronization request returned."
        Write-Host ""
        Write-Host "Server       : $($Result.Server)"
        Write-Host "Requested At : $($Result.RequestedAt.ToString('MM/dd/yyyy hh:mm:ss tt'))"
        Write-Host "Result       : $($Result.Result)"
        Write-Host ""


        # ----------------------------------------------------
        # START LIVE MONITORING ONLY IF REQUEST WAS ACCEPTED
        # ----------------------------------------------------

        if ($Result.WasAccepted) {

            Write-Host "Delta synchronization request accepted."

            $WatchParams = @{
                Credential  = $Credential
                RequestedAt = $Result.RequestedAt
            }

            Watch-EntraDeltaSync @WatchParams
        }
        else {

            Write-Warning "The Delta synchronization request was not accepted."
            Write-Host ""
            Write-Host "Returned Result : $($Result.Result)"
            Write-Host ""
            Write-Host "Synchronization monitoring will not be started."
            Write-Host ""
        }
    }
    catch {

        Write-Host ""
        Write-Error "Unable to start the Delta synchronization."
        Write-Host $_.Exception.Message
        Write-Host ""
    }
}


# ============================================================
# MAIN MENU
# ============================================================

:MainMenu while ($true) {

    Write-Header

    Write-Host "[1] Start Delta Sync"
    Write-Host "[2] Check Sync Status"
    Write-Host "[3] Update Stored Credential"
    Write-Host "[4] Test Connection"
    Write-Host "[5] Exit"
    Write-Host ""

    $Selection = Read-Host "Selection"

    switch ($Selection) {

        "1" {

            $Cred = Get-ValidatedCredential

            if ($null -ne $Cred) {
                Start-EntraDeltaSync -Credential $Cred
            }

            Wait-ForEnter
        }


        "2" {

            $Cred = Get-ValidatedCredential

            if ($null -ne $Cred) {
                Get-EntraSyncStatus -Credential $Cred
            }

            Wait-ForEnter
        }


        "3" {

            Update-EntraSyncCredential | Out-Null

            Wait-ForEnter
        }


        "4" {

            $Cred = Get-EntraSyncCredential

            if ($null -ne $Cred) {

                $Connection = Test-EntraSyncConnection -Credential $Cred

                if ($null -eq $Connection) {
                    Invoke-CredentialRecovery | Out-Null
                }
            }

            Wait-ForEnter
        }


        "5" {

            Write-Host ""
            Write-Host "Exiting Entra Connect Remote Sync Utility."
            Write-Host ""

            break MainMenu
        }


        default {

            Write-Host ""
            Write-Warning "Invalid selection. Please select 1 through 5."

            Start-Sleep -Seconds 1
        }
    }
}