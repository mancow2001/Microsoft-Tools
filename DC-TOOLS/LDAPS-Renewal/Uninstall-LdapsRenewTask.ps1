<#
.SYNOPSIS
    Uninstalls the LDAPS Certificate Renewal solution.

.DESCRIPTION
    Removes the LDAPS certificate renewal solution including:
    - Scheduled task
    - Installation directory (C:\Program Files\LDAPS-Renewal)
    - Optionally removes logs and working files

    Does NOT remove any certificates from the certificate store.

.PARAMETER TaskName
    Name of the scheduled task to remove. Default: "LDAPS Cert Renewal"

.PARAMETER RemoveLogs
    Also remove log files and working directory (C:\ProgramData\LdapsCertRenew)

.PARAMETER Force
    Skip confirmation prompts

.EXAMPLE
    .\Uninstall-LdapsRenewTask.ps1
    # Removes task and scripts, prompts for confirmation

.EXAMPLE
    .\Uninstall-LdapsRenewTask.ps1 -Force
    # Removes task and scripts without prompting

.EXAMPLE
    .\Uninstall-LdapsRenewTask.ps1 -RemoveLogs -Force
    # Complete removal including logs, without prompting

.NOTES
    Version: 1.5.1
    Author: PKI Automation
    Requires: Windows Server 2012 R2+, PowerShell 4.0+, Administrator privileges

    This script does NOT remove certificates. Existing LDAPS certificates
    will remain in the certificate store and continue to function.

    v1.5.1 - Fixed PowerShell 4.0 strict mode compatibility
#>

#Requires -Version 4.0
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TaskName = "LDAPS Cert Renewal",

    [Parameter(Mandatory = $false)]
    [switch]$RemoveLogs,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Constants
$script:InstallPath = Join-Path -Path $env:ProgramFiles -ChildPath "LDAPS-Renewal"
$script:LogPath = "C:\ProgramData\LdapsCertRenew"
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
#endregion

#region Main Uninstall
Write-Host ""
Write-Host "=" * 60
Write-Host "LDAPS Certificate Renewal - Uninstall"
Write-Host "Version: $script:Version"
Write-Host "=" * 60
Write-Host ""

# Check what exists
$taskExists = $null -ne (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)
$installExists = Test-Path -Path $script:InstallPath
$logsExist = Test-Path -Path $script:LogPath

Write-Status "Checking installed components..."
Write-Host ""

if (-not $taskExists -and -not $installExists -and -not $logsExist) {
    Write-Status "No installation found" -Type Warning
    Write-Host ""
    Write-Host "Nothing to uninstall. The following were checked:"
    Write-Host "  - Scheduled task: $TaskName"
    Write-Host "  - Install directory: $script:InstallPath"
    Write-Host "  - Log directory: $script:LogPath"
    exit 0
}

# Display what will be removed
Write-Host "The following components were found:"
Write-Host ""

if ($taskExists) {
    Write-Host "  [x] Scheduled Task: $TaskName" -ForegroundColor Yellow
}
else {
    Write-Host "  [ ] Scheduled Task: $TaskName (not found)"
}

if ($installExists) {
    $files = @(Get-ChildItem -Path $script:InstallPath -Recurse -File -ErrorAction SilentlyContinue)
    Write-Host "  [x] Installation Directory: $script:InstallPath" -ForegroundColor Yellow
    Write-Host "      Contains $($files.Count) file(s)"
}
else {
    Write-Host "  [ ] Installation Directory: $script:InstallPath (not found)"
}

if ($logsExist) {
    $logFiles = @(Get-ChildItem -Path $script:LogPath -Recurse -File -ErrorAction SilentlyContinue)
    $measureResult = $logFiles | Measure-Object -Property Length -Sum
    $logSize = 0
    if ($null -ne $measureResult -and $null -ne $measureResult.Sum) {
        $logSize = $measureResult.Sum
    }
    $logSizeMB = [math]::Round($logSize / 1MB, 2)

    if ($RemoveLogs) {
        Write-Host "  [x] Log Directory: $script:LogPath" -ForegroundColor Yellow
        Write-Host "      Contains $($logFiles.Count) file(s), $logSizeMB MB"
    }
    else {
        Write-Host "  [ ] Log Directory: $script:LogPath (will be preserved)"
        Write-Host "      Use -RemoveLogs to also remove logs"
    }
}
else {
    Write-Host "  [ ] Log Directory: $script:LogPath (not found)"
}

Write-Host ""

# Confirm uninstall
if (-not $Force) {
    Write-Host "This will remove the LDAPS Certificate Renewal solution." -ForegroundColor Yellow
    Write-Host "Existing certificates will NOT be removed." -ForegroundColor Green
    Write-Host ""
    $response = Read-Host "Are you sure you want to uninstall? (Y/N)"
    if ($response -notin @('Y', 'y', 'Yes', 'yes')) {
        Write-Status "Uninstall cancelled" -Type Warning
        exit 0
    }
    Write-Host ""
}

$errors = @()

# Remove scheduled task
if ($taskExists) {
    Write-Status "Removing scheduled task: $TaskName"
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Status "Scheduled task removed" -Type Success
    }
    catch {
        Write-Status "Failed to remove scheduled task: $_" -Type Error
        $errors += "Scheduled task: $_"
    }
}

# Remove installation directory
if ($installExists) {
    Write-Status "Removing installation directory: $script:InstallPath"
    try {
        Remove-Item -Path $script:InstallPath -Recurse -Force
        Write-Status "Installation directory removed" -Type Success
    }
    catch {
        Write-Status "Failed to remove installation directory: $_" -Type Error
        $errors += "Installation directory: $_"
    }
}

# Remove logs if requested
if ($RemoveLogs -and $logsExist) {
    Write-Status "Removing log directory: $script:LogPath"
    try {
        Remove-Item -Path $script:LogPath -Recurse -Force
        Write-Status "Log directory removed" -Type Success
    }
    catch {
        Write-Status "Failed to remove log directory: $_" -Type Error
        $errors += "Log directory: $_"
    }
}

# Summary
Write-Host ""
Write-Host "=" * 60
Write-Host "Uninstall Summary"
Write-Host "=" * 60
Write-Host ""

if ($errors.Count -eq 0) {
    Write-Status "Uninstall completed successfully" -Type Success
    Write-Host ""
    Write-Host "Removed:"
    if ($taskExists) {
        Write-Host "  - Scheduled task: $TaskName"
    }
    if ($installExists) {
        Write-Host "  - Installation directory: $script:InstallPath"
    }
    if ($RemoveLogs -and $logsExist) {
        Write-Host "  - Log directory: $script:LogPath"
    }

    if (-not $RemoveLogs -and $logsExist) {
        Write-Host ""
        Write-Host "Preserved:"
        Write-Host "  - Log directory: $script:LogPath"
        Write-Host "    (run with -RemoveLogs to remove)"
    }

    Write-Host ""
    Write-Host "Note: Existing LDAPS certificates were NOT removed and will"
    Write-Host "continue to function until they expire."

    exit 0
}
else {
    Write-Status "Uninstall completed with errors" -Type Warning
    Write-Host ""
    Write-Host "Errors encountered:"
    foreach ($err in $errors) {
        Write-Host "  - $err" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Some components may need to be removed manually."

    exit 1
}
#endregion
