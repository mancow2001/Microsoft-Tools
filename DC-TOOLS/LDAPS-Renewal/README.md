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

- Windows Server 2016 or later
- PowerShell 5.1 or later
- Administrator/SYSTEM access
- Network connectivity to Enterprise CA
- `certreq.exe` available (built into Windows)

## Files

| File | Purpose |
|------|---------|
| `Renew-LdapsCert.ps1` | Main certificate renewal script |
| `Install-LdapsRenewTask.ps1` | Scheduled task installer |
| `README.md` | This documentation |

## Installation

### Step 1: Deploy Scripts

Copy both `.ps1` files to a secure location on each Domain Controller:

```powershell
# Recommended location
$deployPath = "C:\Scripts\LDAPS-Renewal"
New-Item -Path $deployPath -ItemType Directory -Force

# Copy files (from your deployment source)
Copy-Item -Path ".\Renew-LdapsCert.ps1" -Destination $deployPath
Copy-Item -Path ".\Install-LdapsRenewTask.ps1" -Destination $deployPath
```

### Step 2: Test Manual Execution

Before scheduling, test the script manually with `-WhatIf`:

```powershell
# Navigate to script location
cd C:\Scripts\LDAPS-Renewal

# Test with WhatIf (no changes made)
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -WhatIf

# If successful, run without WhatIf
.\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA"
```

### Step 3: Install Scheduled Task

```powershell
# Basic installation
.\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-CA"

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

## Configuration Reference

### Renew-LdapsCert.ps1 Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-CAConfig` | Yes | - | CA config string (e.g., "CAHOST\CA-NAME") |
| `-TemplateName` | No | LDAPS | Certificate template name |
| `-BaseDomain` | No | - | Additional SAN DNS entry for base domain |
| `-IncludeShortNameSan` | No | $true | Include DC hostname in SAN |
| `-RenewWithinDays` | No | 45 | Days before expiration to trigger renewal |
| `-LogPath` | No | C:\ProgramData\LdapsCertRenew\renew.log | Log file path |
| `-CleanupOld` | No | $false | Remove superseded LDAPS certs |
| `-WhatIf` | No | $false | Preview mode (no changes) |
| `-MinKeySize` | No | 2048 | RSA key size |
| `-HashAlgorithm` | No | sha256 | Hash algorithm (sha256/sha384/sha512) |
| `-VerboseLogging` | No | $false | Enable DEBUG/TRACE level logging for troubleshooting |

### Install-LdapsRenewTask.ps1 Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-CAConfig` | Yes | - | CA config string |
| `-TemplateName` | No | LDAPS | Certificate template name |
| `-BaseDomain` | No | - | Additional SAN DNS entry |
| `-IncludeShortNameSan` | No | $true | Include hostname in SAN |
| `-RenewWithinDays` | No | 45 | Renewal threshold in days |
| `-CleanupOld` | No | $false | Auto-cleanup superseded certs |
| `-VerboseLogging` | No | $false | Enable verbose logging in scheduled runs |
| `-ScriptPath` | No | (same directory) | Path to Renew-LdapsCert.ps1 |
| `-TaskName` | No | LDAPS Cert Renewal | Scheduled task name |
| `-TriggerDay` | No | Sunday | Day of week for weekly trigger |
| `-TriggerTime` | No | 03:15 | Time for trigger (HH:mm) |
| `-RandomDelayMinutes` | No | 30 | Random delay to stagger DCs |
| `-Force` | No | $false | Overwrite existing task |
| `-Uninstall` | No | $false | Remove the scheduled task |

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
[2024-03-15 03:15:22.456] [INFO] ======================================================================
[2024-03-15 03:15:22.456] [INFO] Certificate Discovery
[2024-03-15 03:15:22.456] [INFO] ======================================================================
[2024-03-15 03:15:22.512] [INFO] Searching for LDAPS candidate certificates...
[2024-03-15 03:15:22.523] [INFO] DC FQDN: dc01.contoso.com
[2024-03-15 03:15:22.534] [INFO] Total candidates found: 0
[2024-03-15 03:15:22.545] [INFO] STATE C: No LDAPS candidate certificates found - bootstrap enrollment required
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
[2024-03-15 03:15:22.456] [+0.000s] [INFO] ======================================================================
[2024-03-15 03:15:22.456] [+0.000s] [INFO] LDAPS Certificate Renewal - Started
[2024-03-15 03:15:22.456] [+0.000s] [INFO] ======================================================================
[2024-03-15 03:15:22.457] [+0.001s] [INFO] Script version: 1.1.0
[2024-03-15 03:15:22.458] [+0.002s] [INFO] Verbose logging: True
...
[2024-03-15 03:15:22.512] [+0.056s] [DEBUG] [1/5] Evaluating certificate: ABC123DEF456...
[2024-03-15 03:15:22.513] [+0.057s] [DEBUG]   Certificate Details:
[2024-03-15 03:15:22.514] [+0.058s] [DEBUG]     Thumbprint: ABC123DEF456...
[2024-03-15 03:15:22.515] [+0.059s] [DEBUG]     Subject: CN=dc01.contoso.com
[2024-03-15 03:15:22.516] [+0.060s] [TRACE]     Extensions count: 8
[2024-03-15 03:15:22.517] [+0.061s] [TRACE]       Extension: Enhanced Key Usage (2.5.29.37) Critical=False
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

# Or remove completely
.\Install-LdapsRenewTask.ps1 -Uninstall
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
| `C:\ProgramData\LdapsCertRenew\` | Working directory |
| `C:\ProgramData\LdapsCertRenew\renew.log` | Primary log file |
| `C:\ProgramData\LdapsCertRenew\ldaps_request_*.inf` | Generated INF files (retained for audit) |
| `C:\ProgramData\LdapsCertRenew\ldaps_request_*.req` | Generated CSR files (retained for audit) |
| `C:\ProgramData\LdapsCertRenew\ldaps_cert_*.cer` | Issued certificates (retained for audit) |

## Multi-DC Deployment

For environments with multiple Domain Controllers:

1. **Deploy via Group Policy Preferences** - Copy scripts to all DCs
2. **Stagger Execution** - Use `-RandomDelayMinutes` to prevent CA overload
3. **Monitor Centrally** - Forward logs to SIEM or central logging

Example GPO deployment script:

```powershell
# Run on each DC (via startup script or scheduled task)
$deployPath = "C:\Scripts\LDAPS-Renewal"

if (-not (Test-Path "$deployPath\Renew-LdapsCert.ps1")) {
    # Copy from SYSVOL or network share
    Copy-Item "\\domain.com\SYSVOL\domain.com\scripts\LDAPS-Renewal\*" -Destination $deployPath -Recurse
}

# Install task if not exists
$task = Get-ScheduledTask -TaskName "LDAPS Cert Renewal" -ErrorAction SilentlyContinue
if ($null -eq $task) {
    & "$deployPath\Install-LdapsRenewTask.ps1" -CAConfig "CA01\Contoso-CA" -Force
}
```

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.1.0 | 2024-03-15 | Added verbose logging with DEBUG/TRACE levels, elapsed time tracking, environment discovery |
| 1.0.0 | 2024-03-15 | Initial release |

## License

Internal use only. Modify as needed for your environment.
