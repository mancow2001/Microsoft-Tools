# LDAPS Certificate Renewal - System Administrator Guide

This guide provides step-by-step instructions for deploying, configuring, and maintaining the automated LDAPS certificate renewal solution on Active Directory Domain Controllers.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [Prerequisites Checklist](#prerequisites-checklist)
3. [Certificate Template Setup](#certificate-template-setup)
4. [Single DC Deployment](#single-dc-deployment)
5. [Multi-DC Deployment via GPO](#multi-dc-deployment-via-gpo)
6. [Verification Procedures](#verification-procedures)
7. [Ongoing Maintenance](#ongoing-maintenance)
8. [Monitoring and Alerting](#monitoring-and-alerting)
9. [Troubleshooting Guide](#troubleshooting-guide)
10. [Rollback Procedures](#rollback-procedures)

---

## Quick Start

For experienced administrators, here's the minimal deployment:

```powershell
# 1. Download scripts to a temporary location
# Both Install-LdapsRenewTask.ps1 and Renew-LdapsCert.ps1 must be in the same directory

# 2. Test (preview mode - optional)
.\Renew-LdapsCert.ps1 -WhatIf

# 3. Run installer (deploys to C:\Program Files\LDAPS-Renewal)
.\Install-LdapsRenewTask.ps1

# 4. Verify
Get-ScheduledTask -TaskName "LDAPS Cert Renewal"
Get-ChildItem "C:\Program Files\LDAPS-Renewal"
```

The installer automatically:
- Creates `C:\Program Files\LDAPS-Renewal`
- Copies the renewal script to the installation directory
- Creates a scheduled task pointing to the installed script
- Auto-discovers the Enterprise CA and AD domain name

---

## Prerequisites Checklist

Complete this checklist before deployment:

### Domain Controller Requirements

| Requirement | How to Verify |
|-------------|---------------|
| Windows Server 2012 R2+ | `[System.Environment]::OSVersion` |
| PowerShell 4.0+ | `$PSVersionTable.PSVersion` |
| Administrator access | `whoami /groups` (look for BUILTIN\Administrators) |
| Network access to CA | `Test-NetConnection CA01 -Port 135` |

### Certificate Authority Requirements

| Requirement | Status |
|-------------|--------|
| Enterprise CA deployed | ☐ |
| LDAPS template created | ☐ |
| Template published to CA | ☐ |
| Domain Controllers have Enroll permission | ☐ |
| "Supply in the request" enabled for Subject Name | ☐ |

### Verify CA Accessibility

```powershell
# List available CAs
certutil -config - -ping

# List templates on specific CA
certutil -CATemplates -config "CA01\Contoso-CA"

# Verify LDAPS template exists
certutil -CATemplates -config "CA01\Contoso-CA" | Select-String "LDAPS"
```

---

## Certificate Template Setup

If you don't have an LDAPS certificate template, follow these steps:

### Step 1: Create Template (Certificate Templates Console)

1. Open `certtmpl.msc` on the CA or from RSAT
2. Right-click **Web Server** template → **Duplicate Template**
3. Configure the new template:

**General Tab:**
| Setting | Value |
|---------|-------|
| Template display name | LDAPS |
| Template name | LDAPS |
| Validity period | 1 year (or per policy) |
| Renewal period | 6 weeks |

**Subject Name Tab:**
| Setting | Value |
|---------|-------|
| Supply in the request | ✓ Selected |

**Extensions Tab:**
| Setting | Value |
|---------|-------|
| Application Policies | Server Authentication |

**Security Tab:**
| Principal | Permissions |
|-----------|-------------|
| Domain Controllers | Read, Enroll |

4. Click **OK** to save

### Step 2: Publish Template (CA Console)

1. Open `certsrv.msc` on the CA
2. Expand the CA → Right-click **Certificate Templates**
3. Select **New** → **Certificate Template to Issue**
4. Select **LDAPS** → Click **OK**

### Step 3: Verify Template Publication

```powershell
# Should show LDAPS template
certutil -CATemplates -config "CA01\Contoso-CA" | Select-String "LDAPS"
```

---

## Single DC Deployment

### Step 1: Download Scripts

Download both scripts to a temporary location on the DC (e.g., Desktop or Downloads):
- `Install-LdapsRenewTask.ps1`
- `Renew-LdapsCert.ps1`

Both files must be in the same directory for the installer to work.

```powershell
# Example: Copy from network share to temp location
Copy-Item -Path "\\FileServer\Scripts\LDAPS-Renewal\*.ps1" -Destination "$env:TEMP\"
cd $env:TEMP
```

### Step 2: Test in Preview Mode (Optional)

```powershell
# Test with WhatIf (no changes made)
.\Renew-LdapsCert.ps1 -WhatIf

# Check the log
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 50
```

**Expected output indicators:**
- `AUTO-DISCOVERED CA:` - CA was found automatically
- `DC FQDN:` - Shows the DC's fully qualified name
- `Auto-including AD domain as base domain SAN:` - Domain name detected
- `STATE A/B/C:` - Current certificate state identified

### Step 3: Run Installer

The installer automatically:
- Creates `C:\Program Files\LDAPS-Renewal` directory
- Copies `Renew-LdapsCert.ps1` to the installation directory
- Creates a scheduled task pointing to the installed script

```powershell
# Basic installation (auto-discovers everything)
.\Install-LdapsRenewTask.ps1

# Or with explicit options
.\Install-LdapsRenewTask.ps1 `
    -TriggerDay Sunday `
    -TriggerTime "03:00" `
    -RenewWithinDays 45 `
    -CleanupOld
```

### Step 4: Verify Installation

```powershell
# Check installed files
Get-ChildItem "C:\Program Files\LDAPS-Renewal"

# Check task exists
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Format-List TaskName, State, TaskPath

# Check task configuration - should point to Program Files
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Select-Object -ExpandProperty Actions

# Check next run time
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal" | Select-Object NextRunTime, LastRunTime, LastTaskResult
```

### Step 5: Clean Up Temporary Files

After successful installation, you can remove the downloaded scripts from the temporary location:

```powershell
# The scripts are now installed in Program Files, temp copies can be removed
Remove-Item "$env:TEMP\Install-LdapsRenewTask.ps1" -ErrorAction SilentlyContinue
Remove-Item "$env:TEMP\Renew-LdapsCert.ps1" -ErrorAction SilentlyContinue
```

---

## Multi-DC Deployment via GPO

For environments with multiple Domain Controllers, use Group Policy for consistent deployment.

### Step 1: Prepare SYSVOL Location

```powershell
# Create script folder in SYSVOL
$sysvolPath = "\\$env:USERDNSDOMAIN\SYSVOL\$env:USERDNSDOMAIN\scripts\LDAPS-Renewal"
New-Item -Path $sysvolPath -ItemType Directory -Force

# Copy scripts (from your local copy or download location)
Copy-Item -Path ".\*.ps1" -Destination $sysvolPath

# Verify
Get-ChildItem $sysvolPath
```

### Step 2: Create Group Policy Object

1. Open **Group Policy Management Console** (`gpmc.msc`)
2. Right-click **Domain Controllers** OU → **Create a GPO in this domain, and Link it here**
3. Name: `LDAPS Certificate Auto-Renewal`
4. Right-click the new GPO → **Edit**

### Step 3: Configure File Deployment

Navigate to: **Computer Configuration** → **Preferences** → **Windows Settings** → **Files**

Create two file entries:

**File 1: Renew-LdapsCert.ps1**

| Setting | Value |
|---------|-------|
| Action | Replace |
| Source file | `\\%USERDNSDOMAIN%\SYSVOL\%USERDNSDOMAIN%\scripts\LDAPS-Renewal\Renew-LdapsCert.ps1` |
| Destination file | `C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1` |

### Step 4: Configure Scheduled Task

Navigate to: **Computer Configuration** → **Preferences** → **Control Panel Settings** → **Scheduled Tasks**

Right-click → **New** → **Scheduled Task (At least Windows 7)**

**General Tab:**

| Setting | Value |
|---------|-------|
| Action | Replace |
| Name | LDAPS Cert Renewal |
| User account | NT AUTHORITY\SYSTEM |
| Run whether user is logged on or not | ✓ |
| Run with highest privileges | ✓ |

**Triggers Tab:**

Click **New** and configure:

| Setting | Value |
|---------|-------|
| Begin the task | On a schedule |
| Settings | Weekly |
| Start | 3:00:00 AM |
| Days | Sunday |
| Enabled | ✓ |
| Random delay | 30 minutes |

**Actions Tab:**

Click **New** and configure:

| Setting | Value |
|---------|-------|
| Action | Start a program |
| Program/script | `powershell.exe` |
| Arguments | `-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "& 'C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1' -StartupDelayMaxSeconds 600 -UseHostnameBasedDelay"` |

> **Important:** Use `-Command` (not `-File`) for proper parameter parsing in PowerShell 4.0. The `-StartupDelayMaxSeconds 600 -UseHostnameBasedDelay` parameters stagger execution across DCs to prevent CA overload.

**Settings Tab:**

| Setting | Value |
|---------|-------|
| Allow task to be run on demand | ✓ |
| Run task as soon as possible after a scheduled start is missed | ✓ |
| If the task fails, restart every | 5 minutes |
| Attempt to restart up to | 3 times |
| Stop the task if it runs longer than | 15 minutes |

### Step 5: Apply and Test GPO

```powershell
# On a test DC, force GPO update
gpupdate /force

# Verify GPO applied
gpresult /r /scope:computer | Select-String "LDAPS"

# Verify task was created
Get-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Verify scripts were deployed
Test-Path "C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1"

# Test run the task
Start-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Wait and check results
Start-Sleep -Seconds 30
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 30
```

### Execution Staggering for Large Environments

For environments with many DCs, configure appropriate delays:

| DC Count | StartupDelayMaxSeconds | RandomDelayMinutes |
|----------|------------------------|-------------------|
| 2-5 | 300 | 30 |
| 5-10 | 600 | 30 |
| 10-20 | 900 | 45 |
| 20+ | 1800 | 60 |

---

## Verification Procedures

### Verify Certificate Installation

```powershell
# List all Server Authentication certificates
Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1" -and
    $_.HasPrivateKey
} | Format-Table Thumbprint, Subject, NotAfter, NotBefore

# Check certificate SANs
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*$env:COMPUTERNAME*" } | Select-Object -First 1
$san = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
$san.Format($true)
```

### Verify LDAPS Connectivity

```powershell
# Test LDAPS port locally
Test-NetConnection -ComputerName localhost -Port 636

# Test from another machine
Test-NetConnection -ComputerName DC01.contoso.com -Port 636

# Using OpenSSL (if available)
# openssl s_client -connect DC01.contoso.com:636 -showcerts
```

### Verify Scheduled Task Health

```powershell
# Task status
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Select-Object TaskName, State

# Last run information
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal" | Format-List *

# Interpret LastTaskResult
# 0 = Success
# 1 = Error
# 2 = Pending CA approval
# 267009 = Task is currently running
```

### Check Logs

```powershell
# View recent log entries
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 100

# Search for errors
Select-String -Path "C:\ProgramData\LdapsCertRenew\renew.log" -Pattern "\[ERROR\]"

# Search for warnings
Select-String -Path "C:\ProgramData\LdapsCertRenew\renew.log" -Pattern "\[WARN\]"

# View log from specific date
Select-String -Path "C:\ProgramData\LdapsCertRenew\renew.log" -Pattern "2026-01-30"
```

---

## Ongoing Maintenance

### Weekly Checks

Run these checks weekly or after scheduled task execution:

```powershell
# Quick health check script
$taskInfo = Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1" -and
    $_.Subject -like "*$env:COMPUTERNAME*"
} | Sort-Object NotAfter -Descending | Select-Object -First 1

Write-Host "=== LDAPS Certificate Health Check ===" -ForegroundColor Cyan
Write-Host "Last Task Run: $($taskInfo.LastRunTime)"
Write-Host "Last Result: $($taskInfo.LastTaskResult) $(if($taskInfo.LastTaskResult -eq 0){'(Success)'}else{'(Check logs)'})"
Write-Host "Next Run: $($taskInfo.NextRunTime)"
Write-Host ""
Write-Host "Current Certificate:"
Write-Host "  Thumbprint: $($cert.Thumbprint)"
Write-Host "  Expires: $($cert.NotAfter)"
Write-Host "  Days Remaining: $([math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays))"
```

### Monthly Checks

1. **Review log file size** - Logs rotate automatically at 10MB, but verify:
   ```powershell
   Get-Item "C:\ProgramData\LdapsCertRenew\renew.log" | Select-Object Name, Length, LastWriteTime
   Get-ChildItem "C:\ProgramData\LdapsCertRenew\*.log" | Measure-Object -Property Length -Sum
   ```

2. **Verify CA health** - Ensure the CA is operational:
   ```powershell
   certutil -ping -config "CA01\Contoso-CA"
   ```

3. **Check certificate template** - Ensure template is still published:
   ```powershell
   certutil -CATemplates -config "CA01\Contoso-CA" | Select-String "LDAPS"
   ```

### Updating Scripts

When updating to a new version:

**Single DC (re-run installer):**
```powershell
# Download new versions of Install-LdapsRenewTask.ps1 and Renew-LdapsCert.ps1
# Then re-run the installer with -Force to update

.\Install-LdapsRenewTask.ps1 -Force

# The installer will update the script in Program Files and reconfigure the task
```

**GPO Deployment:**
1. Update scripts in SYSVOL
2. Run `gpupdate /force` on a test DC
3. Verify functionality
4. Allow normal GPO replication to other DCs

### Log Management

Logs are stored at `C:\ProgramData\LdapsCertRenew\` and include:

| File | Purpose | Retention |
|------|---------|-----------|
| `renew.log` | Current log file | Auto-rotates at 10MB |
| `heartbeat.txt` | Last run timestamp/status | Updated each run |
| `renew_YYYYMMDD_HHMMSS.log` | Archived logs | Manual cleanup |
| `ldaps_request_*.inf` | Request config files | Manual cleanup |
| `ldaps_request_*.req` | CSR files | Manual cleanup |
| `ldaps_cert_*.cer` | Issued certificates | Manual cleanup |

**Cleanup old files:**
```powershell
# Remove files older than 90 days
Get-ChildItem "C:\ProgramData\LdapsCertRenew\*" -Exclude "renew.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
    Remove-Item -Force
```

---

## Monitoring and Alerting

### Option 1: Event Log Integration

Add this to a monitoring script that runs after the scheduled task:

```powershell
# Check last run result and write to Event Log
$taskInfo = Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"

if ($taskInfo.LastTaskResult -ne 0) {
    $logContent = Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 50
    Write-EventLog -LogName Application -Source "LDAPS Renewal" -EventId 1001 -EntryType Error -Message "LDAPS certificate renewal failed. Last result: $($taskInfo.LastTaskResult)`n`n$($logContent -join "`n")"
}
```

> **Note:** First create the event source: `New-EventLog -LogName Application -Source "LDAPS Renewal"`

### Option 2: Certificate Expiration Monitoring

Add to your existing monitoring solution:

```powershell
# Check certificate expiration across all DCs
$DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName

foreach ($DC in $DCs) {
    $cert = Invoke-Command -ComputerName $DC -ScriptBlock {
        Get-ChildItem Cert:\LocalMachine\My | Where-Object {
            $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1" -and
            $_.Subject -like "*$env:COMPUTERNAME*"
        } | Sort-Object NotAfter -Descending | Select-Object -First 1
    }

    $daysRemaining = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)

    [PSCustomObject]@{
        DC = $DC
        Thumbprint = $cert.Thumbprint
        Expires = $cert.NotAfter
        DaysRemaining = $daysRemaining
        Status = if ($daysRemaining -lt 30) { "WARNING" } elseif ($daysRemaining -lt 7) { "CRITICAL" } else { "OK" }
    }
} | Format-Table -AutoSize
```

### Option 3: Simple Email Alert

```powershell
# Add to scheduled task or run separately
$threshold = 30  # days
$smtpServer = "mail.contoso.com"
$from = "monitoring@contoso.com"
$to = "sysadmins@contoso.com"

$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1" -and
    $_.Subject -like "*$env:COMPUTERNAME*"
} | Sort-Object NotAfter -Descending | Select-Object -First 1

$daysRemaining = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)

if ($daysRemaining -lt $threshold) {
    $body = @"
WARNING: LDAPS certificate on $env:COMPUTERNAME expires in $daysRemaining days.

Certificate Details:
- Thumbprint: $($cert.Thumbprint)
- Subject: $($cert.Subject)
- Expires: $($cert.NotAfter)

Please investigate the scheduled task 'LDAPS Cert Renewal' or check the log at:
C:\ProgramData\LdapsCertRenew\renew.log
"@

    Send-MailMessage -From $from -To $to -Subject "LDAPS Certificate Expiring on $env:COMPUTERNAME" -Body $body -SmtpServer $smtpServer
}
```

---

## Troubleshooting Guide

### Issue: "No Enterprise CAs discovered"

**Symptoms:** Script fails with "CA auto-discovery failed"

**Diagnosis:**
```powershell
# Check AD connectivity
[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

# Check Configuration partition access
$rootDSE = [ADSI]"LDAP://RootDSE"
$configNC = $rootDSE.configurationNamingContext
Write-Host "Config NC: $configNC"

# Check Enrollment Services
$enrollPath = "LDAP://CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNC"
$enrollServices = [ADSI]$enrollPath
$enrollServices.Children | ForEach-Object { Write-Host $_.cn }
```

**Solutions:**
1. Verify Enterprise CA is installed and configured
2. Check AD replication is healthy
3. Use explicit `-CAConfig` parameter as workaround

### Issue: "Access denied" during enrollment

**Symptoms:** certreq.exe returns permission error

**Diagnosis:**
```powershell
# Check permissions on template
certutil -v -template LDAPS

# Test enrollment permission
certutil -config "CA01\Contoso-CA" -ping
```

**Solutions:**
1. Add Domain Controllers group to template with Enroll permission
2. Verify CA service is running
3. Check firewall rules for RPC (TCP 135) and dynamic ports

### Issue: "Certificate template not found"

**Symptoms:** Template not available for enrollment

**Diagnosis:**
```powershell
# List published templates
certutil -CATemplates -config "CA01\Contoso-CA"
```

**Solutions:**
1. Publish template on CA (see [Certificate Template Setup](#certificate-template-setup))
2. Verify template name matches `-TemplateName` parameter
3. Wait for AD replication if recently published

### Issue: Certificate installed but LDAPS not working

**Symptoms:** Certificate appears in store but port 636 not responding

**Diagnosis:**
```powershell
# Check if LDAPS port is listening
netstat -an | Select-String ":636"

# Check certificate binding
netsh http show sslcert

# Verify certificate has private key
$cert = Get-ChildItem Cert:\LocalMachine\My\<thumbprint>
$cert.HasPrivateKey
```

**Solutions:**
1. **Wait** - NTDS typically picks up new certificates automatically within minutes
2. Force NTDS to re-read certificates:
   ```powershell
   certutil -setreg chain\ChainCacheResyncFiletime @now
   ```
3. As last resort, restart NTDS (requires planning for production):
   ```powershell
   # WARNING: Causes brief AD service interruption
   Restart-Service NTDS -Force
   ```

### Issue: Scheduled task not running

**Symptoms:** Task shows "Ready" but never executes

**Diagnosis:**
```powershell
# Check task state
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Select-Object State, TaskPath

# Check task history
Get-WinEvent -LogName "Microsoft-Windows-TaskScheduler/Operational" -MaxEvents 20 |
    Where-Object { $_.Message -like "*LDAPS*" }

# Verify trigger
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Select-Object -ExpandProperty Triggers
```

**Solutions:**
1. Verify trigger is enabled and configured correctly
2. Check if task is disabled
3. Verify SYSTEM account can execute PowerShell
4. Check execution policy

### Issue: Task runs but no logs are created

**Symptoms:** Task shows success (exit code 0) but no log file or heartbeat

**Diagnosis:**
```powershell
# Check heartbeat file (created even if logging fails)
Test-Path "C:\ProgramData\LdapsCertRenew\heartbeat.txt"
Get-Content "C:\ProgramData\LdapsCertRenew\heartbeat.txt"

# Check Event Log for errors
Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 5 -ErrorAction SilentlyContinue

# Check task arguments (should use -Command not -File)
(Get-ScheduledTask -TaskName "LDAPS Cert Renewal").Actions.Arguments

# Run diagnostics
.\Renew-LdapsCert.ps1 -DiagnoseOnly
```

**Solutions:**
1. **Reinstall with updated scripts** (v1.5.2+) - uses `-Command` instead of `-File` for proper PowerShell 4.0 compatibility
   ```powershell
   .\Install-LdapsRenewTask.ps1 -TemplateName "YourTemplate" -BaseDomain "yourdomain.com" -Force
   ```
2. Verify arguments use `-Command "& '...' ..."` format, not `-File "..." ...`
3. Check that the template name in the task matches an actual published template

### Issue: Request pending CA approval

**Symptoms:** Script returns exit code 2, log shows "pending"

**Diagnosis:**
```powershell
# Check for pending requests on CA
certutil -view -out "RequestID,RequesterName,CommonName,Disposition" -config "CA01\Contoso-CA"
```

**Solutions:**
1. **Option A:** Approve request in CA console
   - Open `certsrv.msc`
   - Navigate to Pending Requests
   - Right-click request → All Tasks → Issue

2. **Option B:** Configure template for auto-approval
   - Edit template → Issuance Requirements
   - Uncheck "CA certificate manager approval"

### Run Diagnostics

Use the built-in diagnostic mode to check configuration without making changes:

```powershell
# Run full diagnostics
.\Renew-LdapsCert.ps1 -DiagnoseOnly

# Check specific template
.\Renew-LdapsCert.ps1 -DiagnoseOnly -TemplateName "LDAPS"
```

Diagnostic mode checks:
- Execution context (SYSTEM vs interactive user)
- Write permissions to log directory
- Domain Controller status
- CA discovery and connectivity
- Template availability on each CA
- Current LDAPS certificates in store
- Scheduled task configuration

### Enable Verbose Logging

For detailed troubleshooting, enable verbose logging:

```powershell
# Run with verbose output
.\Renew-LdapsCert.ps1 -VerboseLogging

# Or update scheduled task arguments to include -VerboseLogging
```

This provides DEBUG and TRACE level information including:
- Full certificate enumeration details
- certreq.exe command output
- CA connectivity tests
- State machine decisions

### Check Heartbeat and Event Log

If logs aren't being created, check the fallback diagnostics:

```powershell
# Heartbeat file is updated every run (even if logging fails)
Get-Content "C:\ProgramData\LdapsCertRenew\heartbeat.txt"

# Event Log fallback (errors written here if file logging fails)
Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 10 -ErrorAction SilentlyContinue
```

---

## Rollback Procedures

### Disable Automatic Renewal

```powershell
# Disable task (preserves configuration)
Disable-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Or uninstall completely
.\Uninstall-LdapsRenewTask.ps1
```

### Remove Problematic Certificate

If a newly enrolled certificate causes issues:

```powershell
# Find the certificate (from logs or by date)
Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.NotBefore -gt (Get-Date).AddDays(-1)
} | Format-Table Thumbprint, Subject, NotBefore

# Remove specific certificate
$thumbprint = "ABC123..."  # Replace with actual thumbprint
Remove-Item "Cert:\LocalMachine\My\$thumbprint"
```

### Restore Previous Certificate

If cleanup removed a working certificate:

**Option 1:** Re-enroll
```powershell
.\Renew-LdapsCert.ps1
```

**Option 2:** Import from backup (if available)
```powershell
Import-PfxCertificate -FilePath "C:\Backup\ldaps_backup.pfx" -CertStoreLocation Cert:\LocalMachine\My -Password (ConvertTo-SecureString -String "password" -AsPlainText -Force)
```

### Complete Removal

To completely remove the solution:

```powershell
# Option 1: Use the uninstall script (recommended)
.\Uninstall-LdapsRenewTask.ps1 -RemoveLogs -Force

# Option 2: Manual removal
# 1. Remove scheduled task
Unregister-ScheduledTask -TaskName "LDAPS Cert Renewal" -Confirm:$false

# 2. Remove installation directory
Remove-Item "C:\Program Files\LDAPS-Renewal" -Recurse -Force

# 3. Remove logs and working files (optional)
Remove-Item "C:\ProgramData\LdapsCertRenew" -Recurse -Force

# Note: Certificates remain in store and will continue to function
```

---

## Appendix: Command Reference

### Renew-LdapsCert.ps1

| Parameter | Description |
|-----------|-------------|
| `-CAConfig` | Explicit CA (e.g., "CA01\Contoso-CA"). Auto-discovers if omitted |
| `-PreferredCA` | Prefer CA matching this name during auto-discovery |
| `-TemplateName` | Certificate template name (default: LDAPS) |
| `-BaseDomain` | Additional SAN entry. Auto-includes AD domain if omitted |
| `-IncludeShortNameSan` | Include hostname in SAN (default: $true) |
| `-RenewWithinDays` | Renewal threshold in days (default: 45) |
| `-CleanupOld` | Remove old certificates after renewal |
| `-WhatIf` | Preview mode - no changes |
| `-VerboseLogging` | Enable DEBUG/TRACE logging |
| `-StartupDelayMaxSeconds` | Max delay before execution (for multi-DC) |
| `-UseHostnameBasedDelay` | Use deterministic delay based on hostname |
| `-DiagnoseOnly` | Run diagnostics without making changes |

### Install-LdapsRenewTask.ps1

| Parameter | Description |
|-----------|-------------|
| `-TaskName` | Scheduled task name (default: "LDAPS Cert Renewal") |
| `-TriggerDay` | Day of week (default: Sunday) |
| `-TriggerTime` | Time in HH:mm format (default: 03:15) |
| `-RandomDelayMinutes` | Task trigger delay (default: 30) |
| `-Force` | Overwrite existing installation without prompting |

*All Renew-LdapsCert.ps1 parameters can also be passed to configure the scheduled task.*

### Uninstall-LdapsRenewTask.ps1

| Parameter | Description |
|-----------|-------------|
| `-TaskName` | Scheduled task name (default: "LDAPS Cert Renewal") |
| `-RemoveLogs` | Also remove log directory |
| `-Force` | Skip confirmation prompts |

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (or no action needed) |
| 1 | Error occurred |
| 2 | Certificate request pending CA approval |

---

## Support

For issues or questions:
1. Check the log file: `C:\ProgramData\LdapsCertRenew\renew.log`
2. Run with `-VerboseLogging` for detailed diagnostics
3. Review the [Troubleshooting Guide](#troubleshooting-guide) above
