# LDAPS Certificate Renewal Solution

Automated LDAPS certificate lifecycle management for Active Directory Domain Controllers.

## Features

- Auto-discovers Enterprise CA from Active Directory
- Auto-includes AD domain name in certificate SAN
- Supports Windows Server 2012 R2 through 2022+
- Runs as SYSTEM via scheduled task (no stored credentials)
- Supports `-WhatIf` for preview/testing

## Files

| File | Purpose |
|------|---------|
| `Renew-LdapsCert.ps1` | Main renewal script |
| `Install-LdapsRenewTask.ps1` | Installer (deploys to Program Files + creates task) |
| `Uninstall-LdapsRenewTask.ps1` | Uninstaller |

## Prerequisites

1. **Certificate Template** with:
   - Server Authentication EKU
   - "Supply in the request" enabled for Subject Name
   - Domain Controllers group has Enroll permission

2. **Domain Controller**: Windows Server 2012 R2+, PowerShell 4.0+

## Installation

```powershell
# Basic (auto-discovers CA and domain)
.\Install-LdapsRenewTask.ps1

# With options
.\Install-LdapsRenewTask.ps1 -TemplateName "LDAPS" -BaseDomain "contoso.com" -CleanupOld
```

## Key Parameters

### Renew-LdapsCert.ps1

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-CAConfig` | (auto) | CA config string (e.g., "CA01\Contoso-CA") |
| `-TemplateName` | LDAPS | Certificate template name |
| `-BaseDomain` | (AD domain) | Additional SAN DNS entry |
| `-RenewWithinDays` | 45 | Days before expiry to renew |
| `-CleanupOld` | false | Remove superseded certs |
| `-WhatIf` | false | Preview mode |
| `-DiagnoseOnly` | false | Run diagnostics only |

### Install-LdapsRenewTask.ps1

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-TriggerDay` | Sunday | Day of week |
| `-TriggerTime` | 03:15 | Time (HH:mm) |
| `-RandomDelayMinutes` | 30 | Stagger execution |
| `-Force` | false | Overwrite existing |

## Verification

```powershell
# Check task
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"

# Check certificate
Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1"
} | Select-Object Thumbprint, Subject, NotAfter

# Check logs
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 50
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| CA not found | Use explicit `-CAConfig "CA01\CA-Name"` |
| Access denied | Grant Domain Controllers Enroll permission on template |
| Template not found | Publish template via `certsrv.msc` |
| No logs created | Reinstall with v1.5.2+ |

### Diagnostic Commands

```powershell
# Run diagnostics
.\Renew-LdapsCert.ps1 -DiagnoseOnly

# Test CA connectivity
certutil -ping -config "CA01\Contoso-CA"

# List available templates
certutil -CATemplates -config "CA01\Contoso-CA"
```

## File Locations

| Path | Purpose |
|------|---------|
| `C:\Program Files\LDAPS-Renewal\` | Installation directory |
| `C:\ProgramData\LdapsCertRenew\renew.log` | Log file |
| `C:\ProgramData\LdapsCertRenew\heartbeat.txt` | Last run status |

## Uninstall

```powershell
.\Uninstall-LdapsRenewTask.ps1 -Force -RemoveLogs
```

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Error |
| 2 | Pending CA approval |

## Version History

| Version | Changes |
|---------|---------|
| 1.5.2 | Fixed RandomDelay format for cross-version compatibility (2012 R2 - 2022) |
| 1.5.1 | Fixed PS 4.0 argument passing; Added `-DiagnoseOnly` |
| 1.5.0 | Installer deploys to Program Files; Separate uninstall script |
