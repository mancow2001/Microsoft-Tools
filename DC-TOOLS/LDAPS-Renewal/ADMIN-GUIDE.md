# LDAPS Certificate Renewal - Quick Start Guide

## Prerequisites

- Windows Server 2012 R2+
- Enterprise CA with LDAPS template published
- Domain Controllers group has **Enroll** permission on template
- Template configured with **"Supply in the request"** for Subject Name

## Installation (Single DC)

```powershell
# 1. Copy both scripts to the DC, then run:
.\Install-LdapsRenewTask.ps1

# 2. Verify
Get-ScheduledTask -TaskName "LDAPS Cert Renewal"
```

The installer auto-discovers the CA and creates a weekly scheduled task.

## Installation (Multi-DC via GPO)

1. Copy scripts to `\\domain\SYSVOL\domain\scripts\LDAPS-Renewal\`
2. Create GPO linked to **Domain Controllers** OU
3. Add **Scheduled Task** (Computer Config → Preferences → Control Panel Settings):
   - Run as: `SYSTEM` | Trigger: Weekly Sunday 3:00 AM | Random delay: 30 min
   - Action: `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1'"`

## Verification

```powershell
# Check task status
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"

# Check certificate
Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1" }

# Check logs
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 50
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| CA not found | Use `-CAConfig "CA01\CA-Name"` explicitly |
| Access denied | Add Domain Controllers to template with Enroll permission |
| Template not found | Publish template on CA via `certsrv.msc` |
| Task runs, no logs | Reinstall with v1.5.2+ (uses `-Command` not `-File`) |

## Uninstall

```powershell
.\Uninstall-LdapsRenewTask.ps1 -Force
```

## File Locations

| Path | Purpose |
|------|---------|
| `C:\Program Files\LDAPS-Renewal\` | Installed script |
| `C:\ProgramData\LdapsCertRenew\renew.log` | Log file |
