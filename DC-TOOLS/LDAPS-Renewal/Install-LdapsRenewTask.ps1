<#
.SYNOPSIS
    Installs or updates the LDAPS Certificate Renewal scheduled task.

.DESCRIPTION
    Creates a scheduled task that runs Renew-LdapsCert.ps1 under SYSTEM context.
    Supports customizable schedule, failure recovery, and idempotent updates.
    Supports CA auto-discovery when -CAConfig is not specified.

.PARAMETER CAConfig
    Optional. CA configuration string (e.g., "CAHOST\CA-NAME").
    If not specified, the script will auto-discover from Active Directory.

.PARAMETER PreferredCA
    When multiple CAs are discovered, prefer CA matching this name (partial match).
    Only used when -CAConfig is not specified.

.PARAMETER TemplateName
    Certificate template name. Default: "LDAPS"

.PARAMETER BaseDomain
    Optional additional SAN DNS entry for base domain

.PARAMETER IncludeShortNameSan
    Include DC hostname in SAN. Default: $true

.PARAMETER RenewWithinDays
    Days before expiration to trigger renewal. Default: 45

.PARAMETER CleanupOld
    Remove superseded certs after successful enrollment

.PARAMETER ScriptPath
    Path to Renew-LdapsCert.ps1. Default: Same directory as this script

.PARAMETER TaskName
    Name of the scheduled task. Default: "LDAPS Cert Renewal"

.PARAMETER TriggerDay
    Day of week for weekly trigger. Default: Sunday

.PARAMETER TriggerTime
    Time for trigger in HH:mm format. Default: 03:15

.PARAMETER RandomDelayMinutes
    Random delay to add to trigger. Default: 30

.PARAMETER Force
    Overwrite existing task without prompting

.PARAMETER Uninstall
    Remove the scheduled task instead of installing

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -BaseDomain "contoso.com"
    # Auto-discovers CA from Active Directory

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-CA" -BaseDomain "contoso.com"
    # Uses explicitly specified CA

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -PreferredCA "Issuing" -BaseDomain "contoso.com"
    # Auto-discovers CA, prefers one with "Issuing" in name

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -TriggerDay Monday -TriggerTime "02:00"

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -Uninstall

.NOTES
    Version: 1.2.0
    Author: PKI Automation
    Requires: Windows Server 2016+, PowerShell 5.1+, Administrator privileges
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(DefaultParameterSetName = 'Install')]
param(
    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [string]$CAConfig,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [string]$PreferredCA,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TemplateName = "LDAPS",

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [string]$BaseDomain,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [bool]$IncludeShortNameSan = $true,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$RenewWithinDays = 45,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [switch]$CleanupOld,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [switch]$VerboseLogging,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ScriptPath,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TaskName = "LDAPS Cert Renewal",

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [string]$TriggerDay = "Sunday",

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidatePattern("^([01]?[0-9]|2[0-3]):[0-5][0-9]$")]
    [string]$TriggerTime = "03:15",

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [ValidateRange(0, 120)]
    [int]$RandomDelayMinutes = 30,

    [Parameter(ParameterSetName = 'Install', Mandatory = $false)]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Uninstall', Mandatory = $true)]
    [switch]$Uninstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper Functions
function Write-Status {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )

    switch ($Type) {
        "Success" { Write-Host "[+] $Message" -ForegroundColor Green }
        "Warning" { Write-Host "[!] $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[-] $Message" -ForegroundColor Red }
        default   { Write-Host "[*] $Message" -ForegroundColor Cyan }
    }
}

function Get-DaysOfWeekFlag {
    param([string]$DayName)

    $dayMap = @{
        "Sunday"    = [System.DayOfWeek]::Sunday
        "Monday"    = [System.DayOfWeek]::Monday
        "Tuesday"   = [System.DayOfWeek]::Tuesday
        "Wednesday" = [System.DayOfWeek]::Wednesday
        "Thursday"  = [System.DayOfWeek]::Thursday
        "Friday"    = [System.DayOfWeek]::Friday
        "Saturday"  = [System.DayOfWeek]::Saturday
    }

    return $dayMap[$DayName]
}
#endregion

#region Uninstall
if ($Uninstall) {
    Write-Status "Uninstalling scheduled task: $TaskName"

    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if ($null -eq $existingTask) {
        Write-Status "Task '$TaskName' does not exist" -Type Warning
        exit 0
    }

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Status "Task '$TaskName' removed successfully" -Type Success
        exit 0
    }
    catch {
        Write-Status "Failed to remove task: $_" -Type Error
        exit 1
    }
}
#endregion

#region Install
Write-Status "Installing LDAPS Certificate Renewal Scheduled Task"
Write-Status "Task Name: $TaskName"

# Determine script path
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Renew-LdapsCert.ps1"
}

# Validate script exists
if (-not (Test-Path -Path $ScriptPath)) {
    Write-Status "Renewal script not found: $ScriptPath" -Type Error
    Write-Status "Ensure Renew-LdapsCert.ps1 is in the same directory or specify -ScriptPath" -Type Error
    exit 1
}

$ScriptPath = (Resolve-Path -Path $ScriptPath).Path
Write-Status "Script path: $ScriptPath"

# Check for existing task
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($null -ne $existingTask -and -not $Force) {
    Write-Status "Task '$TaskName' already exists. Use -Force to update." -Type Warning
    $response = Read-Host "Do you want to update the existing task? (Y/N)"
    if ($response -notin @('Y', 'y', 'Yes', 'yes')) {
        Write-Status "Installation cancelled" -Type Warning
        exit 0
    }
}

# Build script arguments
$scriptArgs = @(
    "-TemplateName `"$TemplateName`""
    "-IncludeShortNameSan `$$IncludeShortNameSan"
    "-RenewWithinDays $RenewWithinDays"
)

# Add CAConfig if explicitly specified
if (-not [string]::IsNullOrWhiteSpace($CAConfig)) {
    $scriptArgs += "-CAConfig `"$CAConfig`""
}

# Add PreferredCA if specified (for auto-discovery)
if (-not [string]::IsNullOrWhiteSpace($PreferredCA)) {
    $scriptArgs += "-PreferredCA `"$PreferredCA`""
}

if (-not [string]::IsNullOrWhiteSpace($BaseDomain)) {
    $scriptArgs += "-BaseDomain `"$BaseDomain`""
}

if ($CleanupOld) {
    $scriptArgs += "-CleanupOld"
}

if ($VerboseLogging) {
    $scriptArgs += "-VerboseLogging"
}

$argumentString = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`" " + ($scriptArgs -join " ")

Write-Status "PowerShell arguments: $argumentString"

# Create action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argumentString

# Create trigger (weekly)
$triggerTimeObj = [DateTime]::ParseExact($TriggerTime, "HH:mm", $null)
$dayOfWeek = Get-DaysOfWeekFlag -DayName $TriggerDay

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dayOfWeek -At $triggerTimeObj

# Add random delay if specified
if ($RandomDelayMinutes -gt 0) {
    $trigger.RandomDelay = "PT${RandomDelayMinutes}M"
}

Write-Status "Trigger: Every $TriggerDay at $TriggerTime (random delay: ${RandomDelayMinutes}m)"

# Create principal (SYSTEM, highest privileges)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Create settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -RestartCount 3 `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
    -MultipleInstances IgnoreNew `
    -Priority 7

# Additional settings not available via New-ScheduledTaskSettingsSet
$settings.DisallowStartOnRemoteAppSession = $false
$settings.RunOnlyIfNetworkAvailable = $true

# Create or update task
try {
    if ($null -ne $existingTask) {
        Write-Status "Updating existing task..."
        Set-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings | Out-Null
    }
    else {
        Write-Status "Creating new task..."
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    }

    # Set task description
    $taskPath = "\$TaskName"
    $existingTask = Get-ScheduledTask -TaskName $TaskName
    $existingTask.Description = "Automated LDAPS certificate renewal for Domain Controller. Manages certificate lifecycle with Enterprise CA."

    Set-ScheduledTask -InputObject $existingTask | Out-Null

    Write-Status "Task '$TaskName' installed successfully" -Type Success
}
catch {
    Write-Status "Failed to create/update task: $_" -Type Error
    exit 1
}

# Display summary
Write-Host ""
Write-Host "=" * 60
Write-Host "Scheduled Task Configuration Summary"
Write-Host "=" * 60
Write-Host "Task Name:         $TaskName"
Write-Host "Run As:            SYSTEM"
Write-Host "Trigger:           Weekly on $TriggerDay at $TriggerTime"
Write-Host "Random Delay:      $RandomDelayMinutes minutes"
Write-Host "Execution Limit:   15 minutes"
Write-Host "Restart on Fail:   Yes (3 attempts, 5 min interval)"
Write-Host ""
Write-Host "Script Parameters:"
Write-Host "  CA Config:       $(if ([string]::IsNullOrWhiteSpace($CAConfig)) { '(auto-discover from AD)' } else { $CAConfig })"
Write-Host "  Preferred CA:    $(if ([string]::IsNullOrWhiteSpace($PreferredCA)) { '(not specified)' } else { $PreferredCA })"
Write-Host "  Template:        $TemplateName"
Write-Host "  Base Domain:     $(if ([string]::IsNullOrWhiteSpace($BaseDomain)) { '(not configured)' } else { $BaseDomain })"
Write-Host "  Include Short:   $IncludeShortNameSan"
Write-Host "  Renew Threshold: $RenewWithinDays days"
Write-Host "  Cleanup Old:     $CleanupOld"
Write-Host "  Verbose Logging: $VerboseLogging"
Write-Host ""
Write-Host "=" * 60

# Offer to run immediately
Write-Host ""
$runNow = Read-Host "Run the task now to verify configuration? (Y/N)"
if ($runNow -in @('Y', 'y', 'Yes', 'yes')) {
    Write-Status "Starting task..."
    try {
        Start-ScheduledTask -TaskName $TaskName
        Write-Status "Task started. Check logs at: C:\ProgramData\LdapsCertRenew\renew.log" -Type Success

        # Wait briefly and check status
        Start-Sleep -Seconds 5
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Status "Last run result: $($taskInfo.LastTaskResult)"

        if ($taskInfo.LastTaskResult -eq 0) {
            Write-Status "Task completed successfully" -Type Success
        }
        elseif ($taskInfo.LastTaskResult -eq 267009) {
            Write-Status "Task is still running..." -Type Info
        }
        else {
            Write-Status "Task may have encountered issues. Check the log file." -Type Warning
        }
    }
    catch {
        Write-Status "Failed to start task: $_" -Type Error
    }
}

Write-Status "Installation complete" -Type Success
exit 0
#endregion
