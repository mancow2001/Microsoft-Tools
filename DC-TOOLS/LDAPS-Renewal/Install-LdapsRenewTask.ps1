<#
.SYNOPSIS
    Installs the LDAPS Certificate Renewal solution and scheduled task.

.DESCRIPTION
    Deploys the LDAPS certificate renewal solution to Program Files and creates
    a scheduled task that runs under SYSTEM context.

    Installation includes:
    - Creating C:\Program Files\LDAPS-Renewal directory
    - Copying Renew-LdapsCert.ps1 to the installation directory
    - Creating a weekly scheduled task to run the renewal script

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
    Additional SAN DNS entry for base domain.
    If not specified, the renewal script auto-includes the AD domain name.

.PARAMETER IncludeShortNameSan
    Include DC hostname in SAN. Default: $true

.PARAMETER RenewWithinDays
    Days before expiration to trigger renewal. Default: 45

.PARAMETER CleanupOld
    Remove superseded certs after successful enrollment

.PARAMETER VerboseLogging
    Enable verbose logging in the renewal script

.PARAMETER TaskName
    Name of the scheduled task. Default: "LDAPS Cert Renewal"

.PARAMETER TriggerDay
    Day of week for weekly trigger. Default: Sunday

.PARAMETER TriggerTime
    Time for trigger in HH:mm format. Default: 03:15

.PARAMETER RandomDelayMinutes
    Random delay to add to trigger. Default: 30

.PARAMETER StartupDelayMaxSeconds
    Maximum startup delay in seconds for the renewal script.
    Helps stagger execution across DCs. Default: 0 (no delay).
    Recommended for multi-DC: 300-900 seconds.

.PARAMETER UseHostnameBasedDelay
    Use deterministic hostname-based delay instead of random.
    Same DC always gets the same delay. Requires StartupDelayMaxSeconds.

.PARAMETER Force
    Overwrite existing installation without prompting

.EXAMPLE
    .\Install-LdapsRenewTask.ps1
    # Auto-discovers CA and AD domain, installs with defaults

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -BaseDomain "contoso.com"
    # Specifies base domain for SAN

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-CA" -Force
    # Uses explicit CA and overwrites existing installation

.EXAMPLE
    .\Install-LdapsRenewTask.ps1 -StartupDelayMaxSeconds 600 -UseHostnameBasedDelay
    # Configures staggered execution for multi-DC environments

.NOTES
    Version: 1.5.1
    Author: PKI Automation
    Requires: Windows Server 2012 R2+, PowerShell 4.0+, Administrator privileges

    Installation Path: C:\Program Files\LDAPS-Renewal
    Log Path: C:\ProgramData\LdapsCertRenew

    Use Uninstall-LdapsRenewTask.ps1 to remove the installation.

    v1.5.1 - Fixed scheduled task argument passing for PowerShell 4.0 compatibility
             (uses -Command instead of -File for proper boolean parsing)
#>

#Requires -Version 4.0
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CAConfig,

    [Parameter(Mandatory = $false)]
    [string]$PreferredCA,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TemplateName = "LDAPS",

    [Parameter(Mandatory = $false)]
    [string]$BaseDomain,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeShortNameSan = $true,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$RenewWithinDays = 45,

    [Parameter(Mandatory = $false)]
    [switch]$CleanupOld,

    [Parameter(Mandatory = $false)]
    [switch]$VerboseLogging,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TaskName = "LDAPS Cert Renewal",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [string]$TriggerDay = "Sunday",

    [Parameter(Mandatory = $false)]
    [ValidatePattern("^([01]?[0-9]|2[0-3]):[0-5][0-9]$")]
    [string]$TriggerTime = "03:15",

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 120)]
    [int]$RandomDelayMinutes = 30,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 3600)]
    [int]$StartupDelayMaxSeconds = 0,

    [Parameter(Mandatory = $false)]
    [switch]$UseHostnameBasedDelay,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Constants
$script:InstallPath = Join-Path -Path $env:ProgramFiles -ChildPath "LDAPS-Renewal"
$script:RenewalScriptName = "Renew-LdapsCert.ps1"
$script:Version = "1.5.1"
#endregion

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

#region Main Installation
Write-Host ""
Write-Host "=" * 60
Write-Host "LDAPS Certificate Renewal - Installation"
Write-Host "Version: $script:Version"
Write-Host "=" * 60
Write-Host ""

Write-Status "Installation path: $script:InstallPath"
Write-Status "Task name: $TaskName"

# Locate source renewal script (must be in same directory as installer)
$sourceScriptPath = Join-Path -Path $PSScriptRoot -ChildPath $script:RenewalScriptName

if (-not (Test-Path -Path $sourceScriptPath)) {
    Write-Status "Renewal script not found: $sourceScriptPath" -Type Error
    Write-Status "Ensure $script:RenewalScriptName is in the same directory as this installer" -Type Error
    exit 1
}

Write-Status "Source script found: $sourceScriptPath"

# Check for existing installation
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
$existingInstall = Test-Path -Path $script:InstallPath

if (($existingTask -or $existingInstall) -and -not $Force) {
    Write-Status "Existing installation detected" -Type Warning
    if ($existingTask) {
        Write-Status "  - Scheduled task '$TaskName' exists"
    }
    if ($existingInstall) {
        Write-Status "  - Installation directory exists: $script:InstallPath"
    }
    Write-Host ""
    $response = Read-Host "Do you want to update the existing installation? (Y/N)"
    if ($response -notin @('Y', 'y', 'Yes', 'yes')) {
        Write-Status "Installation cancelled" -Type Warning
        exit 0
    }
}

# Create installation directory
Write-Status "Creating installation directory..."
if (-not (Test-Path -Path $script:InstallPath)) {
    New-Item -Path $script:InstallPath -ItemType Directory -Force | Out-Null
    Write-Status "Created: $script:InstallPath" -Type Success
}
else {
    Write-Status "Directory already exists: $script:InstallPath"
}

# Copy renewal script to installation directory
$destinationScriptPath = Join-Path -Path $script:InstallPath -ChildPath $script:RenewalScriptName
Write-Status "Copying renewal script..."
Copy-Item -Path $sourceScriptPath -Destination $destinationScriptPath -Force
Write-Status "Installed: $destinationScriptPath" -Type Success

# Verify copy succeeded
if (-not (Test-Path -Path $destinationScriptPath)) {
    Write-Status "Failed to copy script to installation directory" -Type Error
    exit 1
}

# Build script arguments for scheduled task
# Note: Using -Command instead of -File to ensure proper boolean parsing in PS 4.0
$scriptArgs = @(
    "-TemplateName '$TemplateName'"
    "-IncludeShortNameSan:`$$IncludeShortNameSan"
    "-RenewWithinDays $RenewWithinDays"
)

if (-not [string]::IsNullOrWhiteSpace($CAConfig)) {
    $scriptArgs += "-CAConfig '$CAConfig'"
}

if (-not [string]::IsNullOrWhiteSpace($PreferredCA)) {
    $scriptArgs += "-PreferredCA '$PreferredCA'"
}

if (-not [string]::IsNullOrWhiteSpace($BaseDomain)) {
    $scriptArgs += "-BaseDomain '$BaseDomain'"
}

if ($CleanupOld) {
    $scriptArgs += "-CleanupOld"
}

if ($VerboseLogging) {
    $scriptArgs += "-VerboseLogging"
}

if ($StartupDelayMaxSeconds -gt 0) {
    $scriptArgs += "-StartupDelayMaxSeconds $StartupDelayMaxSeconds"
}

if ($UseHostnameBasedDelay) {
    $scriptArgs += "-UseHostnameBasedDelay"
}

# Build command string using -Command instead of -File for proper boolean/variable parsing
$scriptCommand = "& '$destinationScriptPath' " + ($scriptArgs -join " ")
$argumentString = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command `"$scriptCommand`""

Write-Status "Configuring scheduled task..."

# Create scheduled task action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argumentString

# Create trigger (weekly)
$triggerTimeObj = [DateTime]::ParseExact($TriggerTime, "HH:mm", $null)
$dayOfWeek = Get-DaysOfWeekFlag -DayName $TriggerDay
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dayOfWeek -At $triggerTimeObj

# Add random delay if specified (use TimeSpan for 2012 R2 compatibility)
if ($RandomDelayMinutes -gt 0) {
    $trigger.RandomDelay = (New-TimeSpan -Minutes $RandomDelayMinutes)
}

# Create principal (SYSTEM, highest privileges)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Create settings (2012 R2 compatible)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -RestartCount 3 `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
    -MultipleInstances IgnoreNew `
    -Priority 7 `
    -RunOnlyIfNetworkAvailable

# Additional settings - only set if property exists (2012 R2 compatibility)
if ($settings.PSObject.Properties['DisallowStartOnRemoteAppSession']) {
    $settings.DisallowStartOnRemoteAppSession = $false
}

# Create or update task
try {
    if ($null -ne $existingTask) {
        Write-Status "Updating existing scheduled task..."
        Set-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings | Out-Null
    }
    else {
        Write-Status "Creating scheduled task..."
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    }

    # Set task description
    $registeredTask = Get-ScheduledTask -TaskName $TaskName
    $registeredTask.Description = "Automated LDAPS certificate renewal for Domain Controller. Manages certificate lifecycle with Enterprise CA. Installed by LDAPS-Renewal v$script:Version."
    Set-ScheduledTask -InputObject $registeredTask | Out-Null

    Write-Status "Scheduled task configured successfully" -Type Success
}
catch {
    Write-Status "Failed to configure scheduled task: $_" -Type Error
    exit 1
}

# Display summary
Write-Host ""
Write-Host "=" * 60
Write-Host "Installation Summary"
Write-Host "=" * 60
Write-Host ""
Write-Host "Installation Directory:"
Write-Host "  $script:InstallPath"
Write-Host ""
Write-Host "Installed Files:"
Write-Host "  $destinationScriptPath"
Write-Host ""
Write-Host "Scheduled Task:"
Write-Host "  Name:            $TaskName"
Write-Host "  Run As:          SYSTEM"
Write-Host "  Trigger:         Weekly on $TriggerDay at $TriggerTime"
Write-Host "  Random Delay:    $RandomDelayMinutes minutes"
Write-Host "  Execution Limit: 15 minutes"
Write-Host "  Restart on Fail: Yes (3 attempts, 5 min interval)"
Write-Host ""
Write-Host "Script Configuration:"
Write-Host "  CA Config:       $(if ([string]::IsNullOrWhiteSpace($CAConfig)) { '(auto-discover from AD)' } else { $CAConfig })"
Write-Host "  Preferred CA:    $(if ([string]::IsNullOrWhiteSpace($PreferredCA)) { '(not specified)' } else { $PreferredCA })"
Write-Host "  Template:        $TemplateName"
Write-Host "  Base Domain:     $(if ([string]::IsNullOrWhiteSpace($BaseDomain)) { '(auto-detect from AD)' } else { $BaseDomain })"
Write-Host "  Include Short:   $IncludeShortNameSan"
Write-Host "  Renew Threshold: $RenewWithinDays days"
Write-Host "  Cleanup Old:     $CleanupOld"
Write-Host "  Verbose Logging: $VerboseLogging"
Write-Host "  Startup Delay:   $(if ($StartupDelayMaxSeconds -gt 0) { "${StartupDelayMaxSeconds}s $(if ($UseHostnameBasedDelay) { '(hostname-based)' } else { '(random)' })" } else { '(none)' })"
Write-Host ""
Write-Host "Log Location:"
Write-Host "  C:\ProgramData\LdapsCertRenew\renew.log"
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

Write-Host ""
Write-Status "Installation complete" -Type Success
Write-Status "To uninstall, run: Uninstall-LdapsRenewTask.ps1"
exit 0
#endregion
