# LDAPS Certificate Renewal - Administrator Guide

This guide is for Active Directory and Microsoft Certificate Authority administrators responsible for deploying and maintaining the LDAPS certificate renewal solution.

---

## Prerequisites Checklist

Before deploying, ensure the following are in place:

- [ ] Enterprise CA is operational and accessible from all DCs
- [ ] Certificate template is created with correct settings
- [ ] Certificate template is published on the CA
- [ ] Domain Controllers group has Enroll permission on template
- [ ] DCs running Windows Server 2012 R2 or later with PowerShell 4.0+

---

## Certificate Template Configuration

### Creating the LDAPS Template

1. Open **Certificate Templates Console** (`certtmpl.msc`)
2. Duplicate the **Web Server** template
3. Configure the new template:

| Tab | Setting | Value |
|-----|---------|-------|
| **General** | Template display name | LDAPS |
| **General** | Template name | LDAPS |
| **General** | Validity period | 1-2 years (recommended) |
| **General** | Renewal period | 6 weeks |
| **Request Handling** | Purpose | Signature and encryption |
| **Subject Name** | Subject name format | **Supply in the request** |
| **Extensions** | Application Policies | Server Authentication |
| **Security** | Domain Controllers | Enroll (Allow) |
| **Security** | Authenticated Users | Read (Allow) |

### Publishing the Template

1. Open **Certification Authority** console (`certsrv.msc`)
2. Expand your CA → Right-click **Certificate Templates**
3. Select **New** → **Certificate Template to Issue**
4. Select **LDAPS** template and click **OK**

### Verifying Template Availability

```powershell
# From any DC, verify the template is available
certutil -CATemplates -config "CA01\Contoso-Issuing-CA"

# Look for "LDAPS" in the output
```

---

## Installation Methods

### Method 1: Single DC Installation (Interactive)

Run directly on each Domain Controller:

```powershell
# Copy both scripts to the DC
# Run as Administrator:
.\Install-LdapsRenewTask.ps1
```

The installer will:
- Auto-discover the Enterprise CA from Active Directory
- Deploy the script to `C:\Program Files\LDAPS-Renewal\`
- Create a weekly scheduled task
- Offer to run immediately for verification

### Method 2: Multi-DC Installation via GPO

For environments with many DCs, use Group Policy:

1. **Copy scripts to SYSVOL:**
   ```
   \\domain.com\SYSVOL\domain.com\scripts\LDAPS-Renewal\
     ├── Renew-LdapsCert.ps1
     └── Install-LdapsRenewTask.ps1
   ```

2. **Create a GPO** linked to the Domain Controllers OU

3. **Add a Scheduled Task** (Computer Configuration → Preferences → Control Panel Settings → Scheduled Tasks):

   | Setting | Value |
   |---------|-------|
   | Action | Create |
   | Name | LDAPS Cert Renewal |
   | Run as | NT AUTHORITY\SYSTEM |
   | Run with highest privileges | Yes |
   | Trigger | Weekly, Sunday 3:00 AM |
   | Random delay | 30 minutes |
   | Action | Start a program |
   | Program | `powershell.exe` |
   | Arguments | `-NoProfile -ExecutionPolicy Bypass -Command "& 'C:\Program Files\LDAPS-Renewal\Renew-LdapsCert.ps1' -TemplateName 'LDAPS'"` |

4. **Add a Startup Script** (one-time deployment):
   ```powershell
   # GPO Startup Script to deploy the renewal script
   $source = "\\domain.com\SYSVOL\domain.com\scripts\LDAPS-Renewal\Renew-LdapsCert.ps1"
   $destDir = "C:\Program Files\LDAPS-Renewal"
   $dest = "$destDir\Renew-LdapsCert.ps1"

   if (-not (Test-Path $destDir)) {
       New-Item -Path $destDir -ItemType Directory -Force | Out-Null
   }

   Copy-Item -Path $source -Destination $dest -Force
   ```

### Method 3: PowerShell Remoting

```powershell
# Deploy to multiple DCs
$DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName

foreach ($DC in $DCs) {
    Invoke-Command -ComputerName $DC -ScriptBlock {
        # Assumes scripts are accessible via UNC path
        Set-Location "\\domain.com\SYSVOL\domain.com\scripts\LDAPS-Renewal"
        .\Install-LdapsRenewTask.ps1 -Force
    }
}
```

---

## Verification

### Check Scheduled Task Status

```powershell
# View task configuration
Get-ScheduledTask -TaskName "LDAPS Cert Renewal" | Format-List *

# View last run information
Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal"
```

### Check Current Certificate

```powershell
# List LDAPS-capable certificates
Get-ChildItem Cert:\LocalMachine\My | Where-Object {
    $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1"
} | Select-Object Thumbprint, Subject, NotAfter, @{N='DaysLeft';E={($_.NotAfter - (Get-Date)).Days}}
```

### Check Logs

```powershell
# View recent log entries
Get-Content "C:\ProgramData\LdapsCertRenew\renew.log" -Tail 100

# Check heartbeat (confirms script ran)
Get-Content "C:\ProgramData\LdapsCertRenew\heartbeat.txt"
```

### Run Diagnostics

```powershell
# Comprehensive configuration check (makes no changes)
.\Renew-LdapsCert.ps1 -DiagnoseOnly
```

The diagnostic checks:
1. Execution context (SYSTEM vs interactive)
2. File paths and permissions
3. Domain Controller status
4. CA connectivity and discovery
5. Template availability
6. Current LDAPS certificates
7. Scheduled task configuration

---

## Troubleshooting

### Issue: CA Not Found

**Symptoms:** Log shows "No Enterprise CAs discovered" or "CA auto-discovery failed"

**Solutions:**
1. Verify AD connectivity from the DC
2. Check that CA is Enterprise CA (not Standalone)
3. Use explicit CA config:
   ```powershell
   .\Install-LdapsRenewTask.ps1 -CAConfig "CA01\Contoso-Issuing-CA"
   ```

### Issue: Template Not Found

**Symptoms:** Log shows "Template not found" or certreq fails with template error

**Solutions:**
1. Verify template is published on CA:
   ```powershell
   certutil -CATemplates -config "CA01\Contoso-Issuing-CA"
   ```
2. Verify template name matches exactly (case-sensitive):
   ```powershell
   .\Install-LdapsRenewTask.ps1 -TemplateName "YourTemplateName"
   ```
3. Publish template via `certsrv.msc` if missing

### Issue: Access Denied / Permission Denied

**Symptoms:** Enrollment fails with access denied error

**Solutions:**
1. Verify Domain Controllers group has Enroll permission on template
2. Check CA security permissions
3. Run diagnostic to verify SYSTEM context:
   ```powershell
   .\Renew-LdapsCert.ps1 -DiagnoseOnly
   ```

### Issue: Task Runs But No Logs

**Symptoms:** Task shows as completed but no log file created

**Solutions:**
1. Check heartbeat file exists:
   ```powershell
   Test-Path "C:\ProgramData\LdapsCertRenew\heartbeat.txt"
   ```
2. Reinstall with v1.5.2+ (fixes argument passing):
   ```powershell
   .\Install-LdapsRenewTask.ps1 -Force
   ```
3. Check Windows Event Log for errors:
   ```powershell
   Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 10
   ```

### Issue: Certificate Not Renewed

**Symptoms:** Certificate is expiring but not being renewed

**Solutions:**
1. Check renewal threshold (default 45 days):
   ```powershell
   # Certificate must be within this many days of expiry to trigger renewal
   .\Renew-LdapsCert.ps1 -RenewWithinDays 60 -WhatIf
   ```
2. Verify certificate matches renewal criteria (FQDN in SAN)
3. Run with verbose logging:
   ```powershell
   .\Renew-LdapsCert.ps1 -VerboseLogging
   ```

### Issue: Pending CA Approval (Exit Code 2)

**Symptoms:** Log shows enrollment submitted but pending approval

**Solutions:**
1. This is expected if template requires CA manager approval
2. Approve request in CA console (`certsrv.msc` → Pending Requests)
3. Script will retrieve certificate on next run
4. Consider changing template to auto-approve for Domain Controllers

### Issue: Multiple Old Certificates

**Symptoms:** Certificate store has many old LDAPS certificates

**Solutions:**
1. Enable automatic cleanup:
   ```powershell
   .\Install-LdapsRenewTask.ps1 -CleanupOld -Force
   ```
2. Manual cleanup (keeps newest only):
   ```powershell
   .\Renew-LdapsCert.ps1 -CleanupOld
   ```

---

## Diagnostic Commands

### Test CA Connectivity

```powershell
# Ping the CA
certutil -ping -config "CA01\Contoso-Issuing-CA"

# Get CA info
certutil -CAInfo -config "CA01\Contoso-Issuing-CA"
```

### List Available Templates

```powershell
# Templates published on CA
certutil -CATemplates -config "CA01\Contoso-Issuing-CA"

# Template details
certutil -template "LDAPS"
```

### Check Certificate Chain

```powershell
# Verify certificate chain for an LDAPS cert
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -match $env:COMPUTERNAME }
certutil -verify $cert.Thumbprint
```

### Force Manual Enrollment

```powershell
# Preview what would happen
.\Renew-LdapsCert.ps1 -WhatIf -VerboseLogging

# Force enrollment regardless of threshold
.\Renew-LdapsCert.ps1 -RenewWithinDays 365 -VerboseLogging
```

---

## Monitoring

### Log File Location

| File | Purpose |
|------|---------|
| `C:\ProgramData\LdapsCertRenew\renew.log` | Detailed execution log |
| `C:\ProgramData\LdapsCertRenew\heartbeat.txt` | Last run timestamp and parameters |

### Windows Event Log

The script writes to the Application log with source `LDAPS-Renewal` when file logging fails.

```powershell
Get-EventLog -LogName Application -Source "LDAPS-Renewal" -Newest 20
```

### Scheduled Task Monitoring

```powershell
# Check all DCs for task status
$DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
foreach ($DC in $DCs) {
    $taskInfo = Invoke-Command -ComputerName $DC -ScriptBlock {
        Get-ScheduledTaskInfo -TaskName "LDAPS Cert Renewal" -ErrorAction SilentlyContinue
    }
    [PSCustomObject]@{
        DC = $DC
        LastRun = $taskInfo.LastRunTime
        Result = $taskInfo.LastTaskResult
    }
}
```

### Certificate Expiration Report

```powershell
# Report certificates across all DCs
$DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
$report = foreach ($DC in $DCs) {
    Invoke-Command -ComputerName $DC -ScriptBlock {
        Get-ChildItem Cert:\LocalMachine\My | Where-Object {
            $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.1"
        } | Select-Object @{N='DC';E={$env:COMPUTERNAME}}, Thumbprint, Subject, NotAfter,
            @{N='DaysLeft';E={[math]::Floor(($_.NotAfter - (Get-Date)).TotalDays)}}
    }
}
$report | Sort-Object DaysLeft | Format-Table -AutoSize
```

---

## Uninstallation

### Remove from Single DC

```powershell
# Remove task and scripts (keep logs)
.\Uninstall-LdapsRenewTask.ps1

# Complete removal including logs
.\Uninstall-LdapsRenewTask.ps1 -RemoveLogs -Force
```

### Remove via GPO

1. Delete the GPO scheduled task
2. Remove startup script
3. Optionally run cleanup script on each DC

**Note:** Uninstalling does NOT remove certificates. Existing LDAPS certificates will continue to function until they expire.

---

## Best Practices

1. **Test in lab first** - Validate template and configuration before production deployment

2. **Use diagnostics** - Run `-DiagnoseOnly` before deployment to catch issues early

3. **Stagger execution** - In large environments, use startup delays:
   ```powershell
   .\Install-LdapsRenewTask.ps1 -StartupDelayMaxSeconds 600 -UseHostnameBasedDelay
   ```

4. **Enable cleanup** - Prevent certificate store bloat:
   ```powershell
   .\Install-LdapsRenewTask.ps1 -CleanupOld
   ```

5. **Monitor logs** - Periodically review logs for warnings or errors

6. **Set appropriate renewal threshold** - Default 45 days provides ample time for intervention if issues occur

7. **Document CA configuration** - If using explicit CA config, document it for disaster recovery
