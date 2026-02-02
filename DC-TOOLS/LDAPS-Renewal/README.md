# LDAPS Certificate Renewal Solution

Automated LDAPS certificate lifecycle management for Active Directory Domain Controllers using Microsoft Enterprise Certificate Authority.

## Overview

This solution provides automated certificate enrollment and renewal for LDAPS (LDAP over SSL/TLS) on Windows Domain Controllers. It uses `certreq.exe` for certificate operations and runs as a scheduled task under SYSTEM context.

**Version:** 1.5.2

## Files

| File | Version | Description |
|------|---------|-------------|
| `Renew-LdapsCert.ps1` | 1.5.2 | Core renewal script - handles certificate discovery, state evaluation, enrollment, and cleanup |
| `Install-LdapsRenewTask.ps1` | 1.5.2 | Installer - deploys to Program Files and creates scheduled task |
| `Uninstall-LdapsRenewTask.ps1` | 1.5.1 | Uninstaller - removes task and installation files |

## Requirements

- Windows Server 2012 R2 or later
- PowerShell 4.0 or later
- Administrator privileges
- Active Directory domain membership
- Microsoft Enterprise CA with published certificate template

## Architecture

### State Machine

The renewal script operates using a three-state model:

| State | Condition | Action |
|-------|-----------|--------|
| **A** | Valid certificate exists, not within renewal threshold | No action |
| **A** | Valid certificate exists, within renewal threshold | Proactive renewal |
| **B** | Certificate expired or missing required SANs | Immediate enrollment |
| **C** | No LDAPS candidate certificate exists | Bootstrap enrollment |

### Certificate Candidate Criteria

A certificate qualifies as an LDAPS candidate if it meets ALL of these requirements:
1. Located in `LocalMachine\My` certificate store
2. Has a private key
3. Has Server Authentication EKU (OID `1.3.6.1.5.5.7.3.1`)
4. Subject CN or SAN contains the DC's FQDN

### File Locations

| Path | Purpose |
|------|---------|
| `C:\Program Files\LDAPS-Renewal\` | Installation directory (deployed script) |
| `C:\ProgramData\LdapsCertRenew\` | Runtime working directory |
| `C:\ProgramData\LdapsCertRenew\renew.log` | Main log file (rotates at 10MB) |
| `C:\ProgramData\LdapsCertRenew\heartbeat.txt` | Last execution status |
| `C:\ProgramData\LdapsCertRenew\request.inf` | Certificate request configuration |
| `C:\ProgramData\LdapsCertRenew\request.req` | Generated CSR |
| `C:\ProgramData\LdapsCertRenew\response.cer` | CA response certificate |

---

## Renew-LdapsCert.ps1

The core renewal script that manages the certificate lifecycle.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CAConfig` | string | (auto) | CA configuration string (e.g., `"CA01\Contoso-CA"`). Auto-discovers from AD if not specified. |
| `-TemplateName` | string | `LDAPS` | Certificate template name to request |
| `-PreferredCA` | string | - | When multiple CAs exist, prefer one matching this name (partial match) |
| `-BaseDomain` | string | (AD domain) | Additional DNS entry for SAN (auto-uses AD domain if not specified) |
| `-IncludeShortNameSan` | bool | `$true` | Include DC hostname (short name) in SAN |
| `-RenewWithinDays` | int | `45` | Days before expiration to trigger renewal (1-365) |
| `-LogPath` | string | `C:\ProgramData\LdapsCertRenew\renew.log` | Log file location |
| `-CleanupOld` | switch | `$false` | Remove superseded certificates after successful enrollment |
| `-MinKeySize` | int | `2048` | RSA key size (2048, 3072, or 4096) |
| `-HashAlgorithm` | string | `sha256` | Hash algorithm (sha256, sha384, sha512) |
| `-VerboseLogging` | switch | `$false` | Enable DEBUG/TRACE level logging |
| `-StartupDelayMaxSeconds` | int | `0` | Maximum random delay before execution (0-3600) |
| `-UseHostnameBasedDelay` | switch | `$false` | Use deterministic hostname-based delay instead of random |
| `-DiagnoseOnly` | switch | `$false` | Run diagnostics without making changes |
| `-WhatIf` | switch | `$false` | Preview actions without making changes |

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (enrollment completed or no action needed) |
| 1 | Error (enrollment failed or exception occurred) |
| 2 | Pending CA manager approval |

### Key Functions

| Function | Purpose |
|----------|---------|
| `Get-EnterpriseCAs` | Queries AD Configuration partition for Enterprise CAs |
| `Select-CertificateAuthority` | Selects appropriate CA based on preferences and template availability |
| `Get-DcIdentity` | Retrieves DC hostname, FQDN, and domain information |
| `Get-LdapsCandidateCertificates` | Discovers certificates matching LDAPS criteria |
| `Get-CertificateState` | Determines state (A/B/C) and required action |
| `New-CertificateRequestInf` | Generates INF file for `certreq.exe` |
| `Invoke-CertReq` | Executes `certreq.exe` with detailed logging |
| `Invoke-CertificateEnrollment` | Orchestrates the full enrollment workflow |
| `Remove-SupersededCertificates` | Cleans up old certificates after successful enrollment |
| `Invoke-Diagnostics` | Comprehensive configuration verification |

### CA Auto-Discovery

When `-CAConfig` is not specified, the script:
1. Queries `LDAP://CN=Enrollment Services,CN=Public Key Services,CN=Services,<configNC>`
2. Enumerates all `pKIEnrollmentService` objects (Enterprise CAs)
3. Checks each CA for the required template
4. Selects based on: single CA → use it; `-PreferredCA` specified → partial match; otherwise → first with template

### SAN (Subject Alternative Name) Generation

The certificate request includes these DNS SANs:
1. **DC FQDN** (always): `dc01.contoso.com`
2. **DC Hostname** (if `-IncludeShortNameSan:$true`): `dc01`
3. **Base Domain** (if specified or auto-detected): `contoso.com`

### Startup Delay Feature

For multi-DC environments to prevent CA overload:

```powershell
# Random delay 0-600 seconds
.\Renew-LdapsCert.ps1 -StartupDelayMaxSeconds 600

# Deterministic delay based on hostname hash (same delay each run)
.\Renew-LdapsCert.ps1 -StartupDelayMaxSeconds 900 -UseHostnameBasedDelay
```

---

## Install-LdapsRenewTask.ps1

Deploys the solution and creates a scheduled task.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CAConfig` | string | (auto) | CA configuration string |
| `-PreferredCA` | string | - | Preferred CA name for auto-discovery |
| `-TemplateName` | string | `LDAPS` | Certificate template name |
| `-BaseDomain` | string | (auto) | Additional SAN DNS entry |
| `-IncludeShortNameSan` | bool | `$true` | Include hostname in SAN |
| `-RenewWithinDays` | int | `45` | Renewal threshold in days |
| `-CleanupOld` | switch | `$false` | Remove superseded certs |
| `-VerboseLogging` | switch | `$false` | Enable verbose logging |
| `-TaskName` | string | `LDAPS Cert Renewal` | Scheduled task name |
| `-TriggerDay` | string | `Sunday` | Day of week for weekly trigger |
| `-TriggerTime` | string | `03:15` | Time for trigger (HH:mm) |
| `-RandomDelayMinutes` | int | `30` | Random delay added to trigger (0-120) |
| `-StartupDelayMaxSeconds` | int | `0` | Startup delay for renewal script |
| `-UseHostnameBasedDelay` | switch | `$false` | Use hostname-based delay |
| `-Force` | switch | `$false` | Overwrite existing installation |

### Installation Process

1. Validates `Renew-LdapsCert.ps1` exists in same directory
2. Creates `C:\Program Files\LDAPS-Renewal\` directory
3. Copies `Renew-LdapsCert.ps1` to installation directory
4. Creates/updates scheduled task with specified parameters
5. Optionally runs the task immediately for verification

### Scheduled Task Configuration

- **Run As:** SYSTEM
- **Privileges:** Highest
- **Trigger:** Weekly (configurable day/time)
- **Random Delay:** Configurable (default 30 minutes)
- **Execution Limit:** 15 minutes
- **Restart on Failure:** 3 attempts, 5-minute interval
- **Network Required:** Yes

### PowerShell 4.0 Compatibility

The installer uses `-Command` instead of `-File` for the scheduled task action to ensure proper boolean parsing on Windows Server 2012 R2.

---

## Uninstall-LdapsRenewTask.ps1

Removes the installation.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-TaskName` | string | `LDAPS Cert Renewal` | Task name to remove |
| `-RemoveLogs` | switch | `$false` | Also remove log directory |
| `-Force` | switch | `$false` | Skip confirmation prompts |

### Removal Process

1. Removes scheduled task
2. Removes `C:\Program Files\LDAPS-Renewal\` directory
3. Optionally removes `C:\ProgramData\LdapsCertRenew\` (logs)

**Note:** Certificates are NOT removed and will continue to function.

---

## Certificate Template Requirements

The CA certificate template must have:

| Setting | Value |
|---------|-------|
| **Purpose** | Server Authentication |
| **EKU** | Server Authentication (1.3.6.1.5.5.7.3.1) |
| **Subject Name** | "Supply in the request" enabled |
| **Security** | Domain Controllers group has Enroll permission |
| **Publication** | Published on issuing CA |

---

## Usage Examples

### Basic Installation (Auto-discovery)

```powershell
.\Install-LdapsRenewTask.ps1
```

### Installation with Explicit CA

```powershell
.\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-Issuing-CA" -TemplateName "LDAPS" -Force
```

### Multi-DC Environment with Staggered Execution

```powershell
.\Install-LdapsRenewTask.ps1 -StartupDelayMaxSeconds 600 -UseHostnameBasedDelay -CleanupOld
```

### Run Diagnostics

```powershell
.\Renew-LdapsCert.ps1 -DiagnoseOnly
```

### Preview Enrollment (WhatIf)

```powershell
.\Renew-LdapsCert.ps1 -WhatIf
```

### Manual Enrollment with Cleanup

```powershell
.\Renew-LdapsCert.ps1 -CleanupOld -VerboseLogging
```

---

## Version History

| Version | Changes |
|---------|---------|
| 1.5.2 | Fixed RandomDelay format for cross-version compatibility (Server 2012 R2 - 2022) using try/catch fallback |
| 1.5.1 | Fixed PS 4.0 argument passing (uses `-Command` instead of `-File`); Added `-DiagnoseOnly` mode; Fixed strict mode compatibility in uninstaller |
| 1.5.0 | Installer deploys to Program Files; Separate uninstall script; Startup delay feature |

---

## Development Notes

### Error Handling

- Uses `$ErrorActionPreference = "Stop"` for fail-fast behavior
- Comprehensive try/catch blocks with detailed logging
- Falls back to Windows Event Log if file logging fails
- Heartbeat file updated at script start for execution verification

### Logging Levels

| Level | Description |
|-------|-------------|
| INFO | Standard operational messages |
| WARN | Non-fatal issues |
| ERROR | Fatal errors |
| DEBUG | Detailed troubleshooting (requires `-VerboseLogging`) |
| TRACE | Very detailed internals (requires `-VerboseLogging`) |

### Windows Server Compatibility

Tested on:
- Windows Server 2012 R2
- Windows Server 2016
- Windows Server 2019
- Windows Server 2022

Key compatibility considerations:
- Uses `New-TimeSpan` for RandomDelay on 2012 R2, ISO 8601 string on newer versions
- Uses `-Command` instead of `-File` for scheduled task arguments
- Avoids PS 5+ features (e.g., uses `.Path` instead of `.Source` on command objects)
