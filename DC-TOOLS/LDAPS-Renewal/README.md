# LDAPS Certificate Renewal Solution

Automated, production-grade LDAPS certificate lifecycle management for Active Directory Domain Controllers.

## Overview

This solution provides scheduled, unattended renewal of LDAPS certificates on Domain Controllers using Enterprise Microsoft CA. It supports three operational states:

| State | Condition | Action |
|-------|-----------|--------|
| **A** | Valid cert exists, within threshold | Renew |
| **B** | Expired/invalid cert exists | Enroll immediately |
| **C** | No LDAPS cert exists | Bootstrap enrollment |

### Key Features

- **CA Auto-Discovery** - Automatically discovers Enterprise CAs from Active Directory
- **Auto BaseDomain SAN** - Automatically includes AD domain name in certificate SAN
- **Wide OS Support** - Windows Server 2012 R2 through Server 2025
- Runs as SYSTEM via Windows Scheduled Task (no stored credentials)
- Idempotent - safe to run repeatedly
- Supports `-WhatIf` for preview/testing
- Detailed logging for troubleshooting and audit
- Automatic cleanup of superseded certificates (optional)
- Configurable SAN entries (FQDN, hostname, base domain)

## Prerequisites

### Certificate Authority Configuration

1. **Certificate Template Setup**

   Create or configure an LDAPS certificate template with:

   ```
   Template Name: LDAPS (or your custom name)
   Purpose: Server Authentication
   Subject Name: Supply in the request
   Key Usage: Digital Signature, Key Encipherment
   Enhanced Key Usage: Server Authentication (1.3.6.1.5.5.7.3.1)
   Minimum Key Size: 2048-bit RSA
   Validity Period: As per your security policy (typically 1-2 years)
   ```

2. **Enable "Supply in Request" for Subject Name**

   - Open Certificate Templates Console (`certtmpl.msc`)
   - Right-click the template → Properties
   - Go to **Subject Name** tab
   - Select **Supply in the request**
   - Click OK

3. **CA Security Permissions**

   Ensure Domain Controllers can request certificates:
   - Open Certification Authority console (`certsrv.msc`)
   - Right-click the template → Properties → Security
   - Add **Domain Controllers** group with **Enroll** permission

4. **Publish the Template**

   - In CA console, right-click **Certificate Templates**
   - Select **New** → **Certificate Template to Issue**
   - Select your LDAPS template

### Domain Controller Requirements

- Windows Server 2012 R2 or later
- PowerShell 4.0 or later
- Administrator/SYSTEM access
- Network connectivity to Enterprise CA
- `certreq.exe` available (built into Windows)

## Files

| File | Purpose |
|------|---------|
| `Renew-LdapsCert.ps1` | Main certificate renewal script |
| `Install-LdapsRenewTask.ps1` | Installer - deploys to Program Files and creates scheduled task |
| `Uninstall-LdapsRenewTask.ps1` | Uninstaller - removes scheduled task and installed files |
| `README.md` | Technical documentation |
| `ADMIN-GUIDE.md` | System Administrator deployment guide |

## Installation

The installer automatically deploys scripts to `C:\Program Files\LDAPS-Renewal` and creates the scheduled task.

### Step 1: Download Scripts

Download both installer scripts to a temporary location on the Domain Controller:
- `Install-LdapsRenewTask.ps1`
- `Renew-LdapsCert.ps1`

### Step 2: Test Manual Execution (Optional)

Before installing, you can test the renewal script manually with `-WhatIf`:

```powershell
# Test with auto-discovery (recommended)
.\Renew-LdapsCert.ps1 -WhatIf

# Or test with explicit CA
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -WhatIf
```

### Step 3: Run Installer

The installer will:
1. Create `C:\Program Files\LDAPS-Renewal` directory
2. Copy `Renew-LdapsCert.ps1` to the installation directory
3. Create a scheduled task pointing to the installed script

```powershell
# Simplest - auto-discover CA and auto-include AD domain as SAN
.\Install-LdapsRenewTask.ps1

# Auto-discover with preference for specific CA name
.\Install-LdapsRenewTask.ps1 -PreferredCA "Issuing"

# Explicit CA specification
.\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-CA"

# With custom base domain (overrides auto-detection)
.\Install-LdapsRenewTask.ps1 -BaseDomain "contoso.com"

# With all options
.\Install-LdapsRenewTask.ps1 `
    -CAConfig "CA01\Contoso-CA" `
    -TemplateName "LDAPS" `
    -BaseDomain "contoso.com" `
    -IncludeShortNameSan $true `
    -RenewWithinDays 45 `
    -CleanupOld `
    -TriggerDay Sunday `
    -TriggerTime "03:15"
```

### Uninstallation

Use the separate uninstall script to remove the solution:

```powershell
# Standard uninstall (preserves logs)
.\Uninstall-LdapsRenewTask.ps1

# Complete removal including logs
.\Uninstall-LdapsRenewTask.ps1 -RemoveLogs

# Silent uninstall (no prompts)
.\Uninstall-LdapsRenewTask.ps1 -Force
```

## Configuration Reference

### Renew-LdapsCert.ps1 Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-CAConfig` | No | (auto-discover) | CA config string (e.g., "CAHOST\CA-NAME"). If omitted, auto-discovers from AD |
| `-PreferredCA` | No | - | When auto-discovering, prefer CA matching this name (partial match) |
| `-TemplateName` | No | LDAPS | Certificate template name |
| `-BaseDomain` | No | (AD domain) | Additional SAN DNS entry. If omitted, auto-includes the AD domain name |
| `-IncludeShortNameSan` | No | $true | Include DC hostname in SAN |
| `-RenewWithinDays` | No | 45 | Days before expiration to trigger renewal |
| `-LogPath` | No | C:\ProgramData\LdapsCertRenew\renew.log | Log file path |
| `-CleanupOld` | No | $false | Remove superseded LDAPS certs |
| `-WhatIf` | No | $false | Preview mode (no changes) |
| `-MinKeySize` | No | 2048 | RSA key size |
| `-HashAlgorithm` | No | sha256 | Hash algorithm (sha256/sha384/sha512) |
| `-VerboseLogging` | No | $false | Enable DEBUG/TRACE level logging for troubleshooting |
| `-StartupDelayMaxSeconds` | No | 0 | Max random delay (0-3600s) before execution. Recommended: 300-900 for multi-DC |
| `-UseHostnameBasedDelay` | No | $false | Use deterministic hostname-based delay instead of random |
| `-DiagnoseOnly` | No | $false | Run comprehensive diagnostics without making changes |

### Install-LdapsRenewTask.ps1 Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-CAConfig` | No | (auto-discover) | CA config string. If omitted, auto-discovers from AD |
| `-PreferredCA` | No | - | When auto-discovering, prefer CA matching this name |
| `-TemplateName` | No | LDAPS | Certificate template name |
| `-BaseDomain` | No | (AD domain) | Additional SAN DNS entry. If omitted, auto-includes AD domain |
| `-IncludeShortNameSan` | No | $true | Include hostname in SAN |
| `-RenewWithinDays` | No | 45 | Renewal threshold in days |
| `-CleanupOld` | No | $false | Auto-cleanup superseded certs |
| `-VerboseLogging` | No | $false | Enable verbose logging in scheduled runs |
| `-TaskName` | No | LDAPS Cert Renewal | Scheduled task name |
| `-TriggerDay` | No | Sunday | Day of week for weekly trigger |
| `-TriggerTime` | No | 03:15 | Time for trigger (HH:mm) |
| `-RandomDelayMinutes` | No | 30 | Random delay to stagger DCs (task trigger level) |
| `-StartupDelayMaxSeconds` | No | 0 | Script-level startup delay (0-3600s). Recommended: 300-900 |
| `-UseHostnameBasedDelay` | No | $false | Use deterministic hostname-based delay |
| `-Force` | No | $false | Overwrite existing installation without prompting |

### Uninstall-LdapsRenewTask.ps1 Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-TaskName` | No | LDAPS Cert Renewal | Scheduled task name to remove |
| `-RemoveLogs` | No | $false | Also remove log directory (C:\ProgramData\LdapsCertRenew) |
| `-Force` | No | $false | Skip confirmation prompts |

## Automatic Configuration

The script is designed to work with minimal configuration by automatically discovering settings from Active Directory.

### Auto BaseDomain SAN

When `-BaseDomain` is not specified, the script automatically includes the AD domain name (from the DC's domain membership) as a SAN entry. This ensures the certificate works for domain-wide LDAPS queries without manual configuration.

**Example**: On a DC in `contoso.com`, the certificate automatically includes:
- `dc01.contoso.com` (DC FQDN - always included)
- `dc01` (hostname - when `-IncludeShortNameSan $true`)
- `contoso.com` (auto-discovered AD domain)

## CA Auto-Discovery

When `-CAConfig` is not specified, the script automatically discovers Enterprise CAs from Active Directory.

### How It Works

1. **Query AD**: Searches the Configuration partition for `pKIEnrollmentService` objects
   ```
   CN=Enrollment Services,CN=Public Key Services,CN=Services,CN=Configuration,DC=domain,DC=com
   ```

2. **Enumerate CAs**: Lists all Enterprise CAs registered in the forest

3. **Filter by Template**: If `-TemplateName` is specified, prefers CAs that have published that template

4. **Apply Preference**: If `-PreferredCA` is specified, selects CA matching that name

5. **Select CA**: Chooses the best available CA based on criteria

### Selection Logic

| Scenario | Behavior |
|----------|----------|
| Single CA in forest | Uses that CA automatically |
| Multiple CAs, no preference | Uses first available CA with template |
| Multiple CAs, `-PreferredCA` specified | Uses CA matching preference |
| No CAs found | Fails with error |

### Usage Examples

```powershell
# Simplest - auto-discover everything
.\Renew-LdapsCert.ps1

# Auto-discover, prefer CA with "Issuing" in name
.\Renew-LdapsCert.ps1 -PreferredCA "Issuing"

# Auto-discover, but filter by template
.\Renew-LdapsCert.ps1 -TemplateName "LDAPS-Custom"

# Explicit CA (skip auto-discovery)
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-Issuing-CA"
```

### Sample Auto-Discovery Log

```
[2026-01-30 03:15:22.100] [INFO] --------------------------------------------------
[2026-01-30 03:15:22.100] [INFO] CA Configuration Resolution
[2026-01-30 03:15:22.100] [INFO] --------------------------------------------------
[2026-01-30 03:15:22.101] [INFO] No CA specified - initiating auto-discovery...
[2026-01-30 03:15:22.150] [INFO] --------------------------------------------------
[2026-01-30 03:15:22.150] [INFO] Enterprise CA Discovery
[2026-01-30 03:15:22.150] [INFO] --------------------------------------------------
[2026-01-30 03:15:22.151] [INFO] Querying Active Directory for Enterprise CAs...
[2026-01-30 03:15:22.200] [INFO]   Discovered CA: ca01.contoso.com\Contoso-Root-CA
[2026-01-30 03:15:22.201] [INFO]     Template 'LDAPS' published: NO
[2026-01-30 03:15:22.210] [INFO]   Discovered CA: ca02.contoso.com\Contoso-Issuing-CA
[2026-01-30 03:15:22.211] [INFO]     Template 'LDAPS' published: YES
[2026-01-30 03:15:22.212] [INFO] Total Enterprise CAs found: 2
[2026-01-30 03:15:22.213] [INFO] Selecting Certificate Authority...
[2026-01-30 03:15:22.214] [INFO]   Filtered to 1 CA(s) with template 'LDAPS'
[2026-01-30 03:15:22.215] [INFO]   Selected CA: ca02.contoso.com\Contoso-Issuing-CA
[2026-01-30 03:15:22.216] [INFO]
[2026-01-30 03:15:22.217] [INFO] AUTO-DISCOVERED CA: ca02.contoso.com\Contoso-Issuing-CA
```

## Execution Staggering (Multi-DC)

When deploying to multiple Domain Controllers, it's important to prevent all DCs from hitting the CA simultaneously. The solution provides two layers of staggering:

### Layer 1: Task Trigger Delay (RandomDelayMinutes)

This is a Windows Scheduled Task feature that adds a random delay when the task is triggered:

```powershell
.\Install-LdapsRenewTask.ps1 -RandomDelayMinutes 30
```

- Applied by Windows Task Scheduler at trigger time
- Different random delay each run
- Default: 30 minutes

### Layer 2: Script Startup Delay (StartupDelayMaxSeconds)

This is a script-level delay that occurs after the task starts but before certificate operations begin:

```powershell
# Random delay (different each run)
.\Renew-LdapsCert.ps1 -StartupDelayMaxSeconds 600

# Hostname-based delay (consistent per DC)
.\Renew-LdapsCert.ps1 -StartupDelayMaxSeconds 900 -UseHostnameBasedDelay
```

### Delay Types

| Type | Parameter | Behavior | Best For |
|------|-----------|----------|----------|
| **Random** | `-StartupDelayMaxSeconds 600` | Different delay each run (0 to max) | General load distribution |
| **Hostname-based** | `-StartupDelayMaxSeconds 900 -UseHostnameBasedDelay` | Same delay for same DC every run | Predictable scheduling |

### How Hostname-Based Delay Works

The hostname-based delay uses a SHA256 hash of the computer name to generate a deterministic delay:

1. Takes the DC hostname (e.g., `DC01`)
2. Computes SHA256 hash
3. Converts first 4 bytes to integer
4. Applies modulo to fit within max delay range

This ensures:
- Same DC always gets the same delay
- Different DCs get evenly distributed delays
- Delays are predictable for troubleshooting

### Recommended Configuration

For environments with multiple DCs:

```powershell
# Install with both delay layers
.\Install-LdapsRenewTask.ps1 `
    -BaseDomain "contoso.com" `
    -RandomDelayMinutes 30 `
    -StartupDelayMaxSeconds 600 `
    -UseHostnameBasedDelay
```

| DC Count | Recommended StartupDelayMaxSeconds |
|----------|-----------------------------------|
| 2-5 DCs | 300 (5 minutes) |
| 5-10 DCs | 600 (10 minutes) |
| 10-20 DCs | 900 (15 minutes) |
| 20+ DCs | 1800 (30 minutes) |

### Sample Delay Log Output

```
[2026-01-30 03:15:22.100] [INFO] ======================================================================
[2026-01-30 03:15:22.100] [INFO] Startup Delay
[2026-01-30 03:15:22.100] [INFO] ======================================================================
[2026-01-30 03:15:22.101] [INFO] Delay type: Hostname-based (deterministic)
[2026-01-30 03:15:22.102] [INFO] Hostname: DC01
[2026-01-30 03:15:22.103] [INFO] Computed delay: 247 seconds (of 600 max)
[2026-01-30 03:15:22.104] [INFO] Waiting 247 seconds before proceeding...
[2026-01-30 03:15:22.105] [INFO] Estimated start time: 2026-01-30 03:19:29
[2026-01-30 03:19:29.200] [INFO] Startup delay completed
```

## Initial Bootstrap Scenario

When no LDAPS certificate exists (State C), the script performs a full bootstrap:

1. **Detection**: Script scans `Cert:\LocalMachine\My` for candidates
2. **No Match**: No certificates found with Server Auth EKU + DC FQDN
3. **Bootstrap**: Generates new PKCS#10 request with all configured SANs
4. **Submit**: Sends request to Enterprise CA using `certreq.exe`
5. **Install**: Accepts issued certificate into LocalMachine\My
6. **Verify**: Confirms all requirements are met

Example bootstrap log:

```
[2026-01-30 03:15:22.456] [INFO] ======================================================================
[2026-01-30 03:15:22.456] [INFO] Certificate Discovery
[2026-01-30 03:15:22.456] [INFO] ======================================================================
[2026-01-30 03:15:22.512] [INFO] Searching for LDAPS candidate certificates...
[2026-01-30 03:15:22.523] [INFO] DC FQDN: dc01.contoso.com
[2026-01-30 03:15:22.534] [INFO] Total candidates found: 0
[2026-01-30 03:15:22.545] [INFO] STATE C: No LDAPS candidate certificates found - bootstrap enrollment required
```

## Validation Steps

### Verify Scheduled Task

```powershell
# Check task exists and configuration
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Format-List *

# Check task history
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"

# View last run result
(Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal").LastTaskResult
# 0 = Success, 1 = Error, 2 = Pending CA approval
```

### Verify Certificate Installation

```powershell
# List LDAPS candidate certificates
Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.HasPrivateKey -and
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1"
} | Select-Object Thumbprint, Subject, NotAfter

# Check specific certificate SANs
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq "YOUR_THUMBPRINT" }
$san = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.17" }
$san.Format($true)
```

### Verify LDAPS Connectivity

```powershell
# From a remote machine
Test-NetConnection -ComputerName dc01.contoso.com -Port 636

# Using OpenSSL (if available)
openssl s_client -connect dc01.contoso.com:636 -showcerts
```

### Review Logs

```powershell
# View recent log entries
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 100

# Search for errors
Select-String -Path "C:\ProgramData\LdapsCertRenew\renew.log" -Pattern "\[ERROR\]"

# Search for warnings
Select-String -Path "C:\ProgramData\LdapsCertRenew\renew.log" -Pattern "\[WARN\]"
```

## Troubleshooting

### Common Issues

#### "Certificate template not found"

**Cause**: Template not published or wrong name

**Solution**:
```powershell
# List available templates on CA
certutil -CATemplates -config "CA01\Contoso-CA"
```

#### "Access denied" during enrollment

**Cause**: DC lacks Enroll permission on template

**Solution**:
1. Open CA console (`certsrv.msc`)
2. Navigate to Certificate Templates
3. Right-click template → Properties → Security
4. Ensure Domain Controllers has **Enroll** permission

#### "Request is pending"

**Cause**: CA requires manager approval

**Solution**:
1. Check CA for pending requests
2. Approve the request manually, OR
3. Reconfigure template to auto-issue without approval

#### Certificate installed but LDAPS not working

**Cause**: LDAPS service may need restart or certificate binding issue

**Note**: This script intentionally does NOT restart services. If needed:
```powershell
# Force NTDS to pick up new certificate (requires downtime planning)
# Option 1: Wait for automatic detection
# Option 2: Use certutil to select the certificate
certutil -setreg chain\ChainCacheResyncFiletime @now
```

#### Missing SAN entries

**Cause**: Template configured to override requestor SANs

**Solution**:
1. Open Certificate Templates Console
2. Edit template → Subject Name tab
3. Ensure "Supply in the request" is selected
4. Uncheck any options that override SANs

#### Scheduled task runs but no logs created

**Cause**: Script fails silently under SYSTEM context (PowerShell 4.0 strict mode issue)

**Diagnosis**:
```powershell
# Check heartbeat file (created even if logging fails)
Get-Content "C:\ProgramData\LdapsCertRenew\heartbeat.txt"

# Check Event Log for errors
Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 5 -ErrorAction SilentlyContinue

# Run diagnostics
.\Renew-LdapsCert.ps1 -DiagnoseOnly
```

**Solution**:
1. Ensure you're using version 1.5.2 or later
2. Reinstall with the updated `Install-LdapsRenewTask.ps1` (uses `-Command` instead of `-File`)
3. Verify the scheduled task arguments use `-Command` not `-File`:
   ```powershell
   (Get-ScheduledTask -TaskName "LDAPS Cert Renewal").Actions.Arguments
   ```

### Verbose Logging Mode

The script supports comprehensive verbose logging for troubleshooting. Enable it with `-VerboseLogging`:

```powershell
# Run with full verbose logging
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -VerboseLogging

# Combine with WhatIf for safe testing
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -VerboseLogging -WhatIf
```

#### Log Levels

| Level | Description | When Shown |
|-------|-------------|------------|
| INFO | Normal operations and status | Always |
| WARN | Warnings and non-critical issues | Always |
| ERROR | Errors and failures | Always |
| DEBUG | Detailed operation information | With `-VerboseLogging` |
| TRACE | Low-level diagnostic data | With `-VerboseLogging` |

#### What Verbose Logging Includes

- **Environment Discovery**: OS version, PowerShell version, domain info, CA connectivity tests
- **Certificate Enumeration**: Full details of every certificate in the store (thumbprint, subject, issuer, SANs, EKUs, extensions)
- **State Machine**: Detailed decision logic for each state transition
- **INF File Contents**: Complete generated INF file content
- **certreq.exe Output**: Full stdout/stderr from all certreq operations
- **CSR/Certificate Dumps**: Uses `certutil -dump` to show request and certificate details
- **Timing Information**: Elapsed time for each operation with `[+X.XXXs]` prefix
- **Verification Steps**: Detailed pass/fail for each compliance check

#### Sample Verbose Log Output

```
[2026-01-30 03:15:22.456] [+0.000s] [INFO] ======================================================================
[2026-01-30 03:15:22.456] [+0.000s] [INFO] LDAPS Certificate Renewal - Started
[2026-01-30 03:15:22.456] [+0.000s] [INFO] ======================================================================
[2026-01-30 03:15:22.457] [+0.001s] [INFO] Script version: 1.5.2
[2026-01-30 03:15:22.458] [+0.002s] [INFO] Verbose logging: True
...
[2026-01-30 03:15:22.512] [+0.056s] [DEBUG] [1/5] Evaluating certificate: ABC123DEF456...
[2026-01-30 03:15:22.513] [+0.057s] [DEBUG]   Certificate Details:
[2026-01-30 03:15:22.514] [+0.058s] [DEBUG]     Thumbprint: ABC123DEF456...
[2026-01-30 03:15:22.515] [+0.059s] [DEBUG]     Subject: CN=dc01.contoso.com
[2026-01-30 03:15:22.516] [+0.060s] [TRACE]     Extensions count: 8
[2026-01-30 03:15:22.517] [+0.061s] [TRACE]       Extension: Enhanced Key Usage (2.5.29.37) Critical=False
```

### Diagnostic Mode

The script includes a comprehensive diagnostic mode that checks configuration without making changes:

```powershell
# Run full diagnostics
.\Renew-LdapsCert.ps1 -DiagnoseOnly

# Check specific template availability
.\Renew-LdapsCert.ps1 -DiagnoseOnly -TemplateName "LDAPS"
```

Diagnostic mode checks:
- Execution context (SYSTEM vs interactive)
- Path and write permissions
- Domain Controller status
- CA discovery and connectivity
- Template availability on each CA
- Current LDAPS certificates
- Scheduled task configuration

This is especially useful for troubleshooting scheduled task issues where no logs are created.

### Heartbeat File

The script creates a heartbeat file (`C:\ProgramData\LdapsCertRenew\heartbeat.txt`) on each run, even if logging fails. Check this file to verify the script is executing:

```powershell
Get-Content "C:\ProgramData\LdapsCertRenew\heartbeat.txt"
```

### Event Log Fallback

If file logging fails, errors are written to the Windows Application Event Log under source "LDAPS-Renewal":

```powershell
Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 10
```

### Debug Mode (Additional Tools)

For additional troubleshooting beyond verbose logging:

```powershell
# Check generated INF file
Get-Content "C:\ProgramData\LdapsCertRenew\ldaps_request_*.inf"

# Check CSR content
certutil -dump "C:\ProgramData\LdapsCertRenew\ldaps_request_*.req"

# Verify certificate details
certutil -dump "C:\ProgramData\LdapsCertRenew\ldaps_cert_*.cer"

# Check certificate in store
certutil -v -store My <thumbprint>

# Test CA connectivity
certutil -ping -config "CA01\Contoso-CA"

# List available templates
certutil -CATemplates -config "CA01\Contoso-CA"
```

## Rollback Procedures

### Disable Scheduled Task

```powershell
# Disable task (keeps configuration)
Disable-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Or remove completely using the uninstall script
.\Uninstall-LdapsRenewTask.ps1 -Force
```

### Remove Newly Installed Certificate

If a new certificate causes issues:

```powershell
# Find the thumbprint from logs or output
$thumbprint = "ABC123..."  # From script output

# Remove the certificate
Remove-Item "Cert:\LocalMachine\My\$thumbprint"

# Re-enable previous certificate (if still present)
# LDAPS will automatically use the remaining valid cert
```

### Restore Previous Certificate

If the old certificate was cleaned up:

1. Re-issue from CA using backup (if available)
2. Import from backup:
   ```powershell
   # If you have a PFX backup
   Import-PfxCertificate -FilePath "backup.pfx" -CertStoreLocation Cert:\LocalMachine\My
   ```
3. Or run the script again to enroll a new certificate

## Security Considerations

- **No Credentials Stored**: Script runs as SYSTEM, uses machine identity
- **Non-Exportable Keys**: Private keys marked as non-exportable
- **Audit Logging**: All operations logged with timestamps
- **Minimal Permissions**: Only accesses LocalMachine\My store
- **No Service Restart**: Script does NOT restart AD DS or other services
- **Template Isolation**: Uses dedicated LDAPS template, won't affect other certs

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (or no action needed) |
| 1 | Error occurred |
| 2 | Certificate request pending CA approval |

## File Locations

| Path | Purpose |
|------|---------|
| `C:\Program Files\LDAPS-Renewal\` | Installation directory |
| `C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1` | Installed renewal script |
| `C:\ProgramData\LdapsCertRenew\` | Working directory for logs and temp files |
| `C:\ProgramData\LdapsCertRenew\renew.log` | Primary log file |
| `C:\ProgramData\LdapsCertRenew\heartbeat.txt` | Heartbeat file (updated each run, useful for diagnostics) |
| `C:\ProgramData\LdapsCertRenew\ldaps_request_*.inf` | Generated INF files (retained for audit) |
| `C:\ProgramData\LdapsCertRenew\ldaps_request_*.req` | Generated CSR files (retained for audit) |
| `C:\ProgramData\LdapsCertRenew\ldaps_cert_*.cer` | Issued certificates (retained for audit) |

## Multi-DC Deployment via Group Policy

This section provides detailed instructions for deploying the LDAPS certificate renewal solution to all Domain Controllers using Group Policy.

### Overview

This method uses Group Policy Preferences to deploy scripts and create a scheduled task directly on all Domain Controllers. It ensures consistent deployment with automatic remediation if scripts are removed.

### Prerequisites

1. **SYSVOL Location**: Store scripts in SYSVOL for replication
   ```
   \\domain.com\SYSVOL\domain.com\scripts\LDAPS-Renewal\
   ├── Renew-LdapsCert.ps1
   └── Install-LdapsRenewTask.ps1
   ```

2. **GPO Targeting**: Create a GPO linked to the **Domain Controllers** OU

3. **Permissions**: The GPO must allow Domain Controllers to read SYSVOL

---

### Step 1: Create the GPO

```
1. Open Group Policy Management Console (gpmc.msc)
2. Right-click "Domain Controllers" OU → Create a GPO
3. Name it: "LDAPS Certificate Auto-Renewal"
4. Right-click the GPO → Edit
```

### Step 2: Deploy Scripts via File Preferences

```
Navigate to:
  Computer Configuration
    → Preferences
      → Windows Settings
        → Files

Create two file entries:
```

**File 1: Renew-LdapsCert.ps1**
| Setting | Value |
|---------|-------|
| Action | Replace |
| Source | `\\%USERDNSDOMAIN%\SYSVOL\%USERDNSDOMAIN%\scripts\LDAPS-Renewal\Renew-LdapsCert.ps1` |
| Destination | `C:\Scripts\LDAPS-Renewal\Renew-LdapsCert.ps1` |

**File 2: Install-LdapsRenewTask.ps1**
| Setting | Value |
|---------|-------|
| Action | Replace |
| Source | `\\%USERDNSDOMAIN%\SYSVOL\%USERDNSDOMAIN%\scripts\LDAPS-Renewal\Install-LdapsRenewTask.ps1` |
| Destination | `C:\Scripts\LDAPS-Renewal\Install-LdapsRenewTask.ps1` |

### Step 3: Create Scheduled Task via GPO Preferences

```
Navigate to:
  Computer Configuration
    → Preferences
      → Control Panel Settings
        → Scheduled Tasks

Right-click → New → Scheduled Task (At least Windows 7)
```

**General Tab:**
| Setting | Value |
|---------|-------|
| Action | Replace |
| Name | LDAPS Cert Renewal |
| User account | NT AUTHORITY\SYSTEM |
| Run whether user is logged on or not | ✓ |
| Run with highest privileges | ✓ |
| Configure for | Windows Server 2012 R2 (or later) |

**Triggers Tab:**
| Setting | Value |
|---------|-------|
| Begin the task | On a schedule |
| Settings | Weekly |
| Start | 3:15:00 AM |
| Days | Sunday (or your preference) |
| Enabled | ✓ |

Click "Advanced settings":
| Setting | Value |
|---------|-------|
| Random delay | 30 minutes |
| Enabled | ✓ |

**Actions Tab:**
| Setting | Value |
|---------|-------|
| Action | Start a program |
| Program/script | `powershell.exe` |
| Arguments | `-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "& 'C:\Scripts\LDAPS-Renewal\Renew-LdapsCert.ps1' -TemplateName 'LDAPS' -BaseDomain 'contoso.com'"` |

> **Note**: Use `-Command` (not `-File`) for proper parameter parsing in PowerShell 4.0. Omit `-CAConfig` to use auto-discovery, or specify explicitly if needed.

**Settings Tab:**
| Setting | Value |
|---------|-------|
| Allow task to be run on demand | ✓ |
| Run task as soon as possible after a scheduled start is missed | ✓ |
| If the task fails, restart every | 5 minutes |
| Attempt to restart up to | 3 times |
| Stop the task if it runs longer than | 15 minutes |
| If the running task does not end when requested, force it to stop | ✓ |

### Step 4: Apply and Test

```powershell
# Force GPO update on a test DC
gpupdate /force

# Verify task was created
Get-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Run task manually to test
Start-ScheduledTask -TaskName "LDAPS Cert Renewal"

# Check results
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 50
```

---

### Best Practices

#### 1. Stagger Execution Across DCs

Use multiple layers of delay to prevent all DCs hitting the CA simultaneously:

```powershell
# Task trigger delay (Windows Scheduler)
-RandomDelayMinutes 30

# Script startup delay (recommended for large environments)
-StartupDelayMaxSeconds 600 -UseHostnameBasedDelay
```

This provides two levels of staggering:
- Task trigger: 0-30 minutes random delay when task fires
- Script startup: 0-600 seconds deterministic delay based on hostname

For GPO Scheduled Task, add the startup delay parameters to the Arguments:
```
-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\LDAPS-Renewal\Renew-LdapsCert.ps1" -TemplateName "LDAPS" -BaseDomain "contoso.com" -StartupDelayMaxSeconds 600 -UseHostnameBasedDelay
```

#### 2. Use WMI Filtering (Optional)

To target only Domain Controllers:

```
WMI Filter: SELECT * FROM Win32_ComputerSystem WHERE DomainRole = 4 OR DomainRole = 5
```

- DomainRole 4 = Backup Domain Controller
- DomainRole 5 = Primary Domain Controller

#### 3. Security Filtering

Ensure only Domain Controllers can read and apply the GPO:

```
Security Filtering:
  - Remove "Authenticated Users"
  - Add "Domain Controllers" with Read and Apply permissions
```

#### 4. Monitoring and Alerting

Forward logs to central monitoring:

```powershell
# Example: Forward to Windows Event Log
$log = Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 100
if ($log -match "\[ERROR\]") {
    Write-EventLog -LogName Application -Source "LDAPS Renewal" -EventId 1001 -EntryType Error -Message ($log -join "`n")
}
```

#### 5. Validate GPO Application

```powershell
# Check GPO application on a DC
gpresult /r /scope:computer

# Verify scheduled task
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Format-List *

# Check recent task runs
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"
```

---

### Troubleshooting GPO Deployment

| Issue | Cause | Solution |
|-------|-------|----------|
| Task not created | GPO not applied | Run `gpupdate /force`, check `gpresult /r` |
| Scripts not copied | SYSVOL access issue | Verify DFS replication, check permissions |
| Task runs but fails | Script path wrong | Verify `C:\Scripts\LDAPS-Renewal\` exists |
| CA not found | Auto-discovery failed | Check AD connectivity, use explicit `-CAConfig` |
| Permission denied | SYSTEM can't access CA | Verify Domain Controllers have Enroll permission |

**Debug GPO Issues:**
```powershell
# Check GPO application
gpresult /h C:\temp\gpresult.html
Start-Process C:\temp\gpresult.html

# Check event logs for GPO errors
Get-WinEvent -LogName "Microsoft-Windows-GroupPolicy/Operational" -MaxEvents 50

# Verify SYSVOL accessibility
Test-Path "\\$env:USERDNSDOMAIN\SYSVOL\$env:USERDNSDOMAIN\scripts\LDAPS-Renewal\Renew-LdapsCert.ps1"
```

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.5.2 | 2026-01 | Fixed nullable datetime parameter for bootstrap scenarios; Additional PS 4.0 strict mode fixes |
| 1.5.1 | 2026-01 | Fixed scheduled task argument passing (uses `-Command` instead of `-File` for PS 4.0); Added `-DiagnoseOnly` parameter; Added heartbeat file and Event Log fallback for troubleshooting |
| 1.5.0 | 2025-01 | Installer now deploys to Program Files; Separate Uninstall script; Improved installation experience |
| 1.4.0 | 2025-01 | Windows Server 2012 R2 compatibility; Auto-include AD domain as BaseDomain SAN when not specified |
| 1.3.0 | 2024-03 | Added execution staggering with `-StartupDelayMaxSeconds` and `-UseHostnameBasedDelay` parameters |
| 1.2.0 | 2024-03 | Added CA auto-discovery from Active Directory, `-PreferredCA` parameter |
| 1.1.0 | 2024-03 | Added verbose logging with DEBUG/TRACE levels, elapsed time tracking, environment discovery |
| 1.0.0 | 2024-03 | Initial release |

## License

Internal use only. Modify as needed for your environment.
