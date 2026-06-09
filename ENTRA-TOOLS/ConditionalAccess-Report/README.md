# Get-ConditionalAccessReport.ps1

Generates a comprehensive, multi-format report of every Conditional Access (CA) policy in an Entra ID tenant, the Enterprise Applications they target, the named locations they reference, and the relationships between them.

This is intended for Entra ID administrators, security engineers, and auditors who need to:

- Inventory the full Conditional Access posture of a tenant.
- Identify which Enterprise Applications are governed by which policies (and which are not).
- Hand off a readable artifact to compliance/audit (Excel, HTML) or to a downstream pipeline (CSV, JSON).
- Detect policy sprawl, gaps, and overlapping scope.

---

## What it produces

Depending on the chosen export format, the script emits a timestamped report containing:

| Section | Contents |
|---------|----------|
| **Summary** | Counts of total, enabled, disabled, and report-only policies; named locations; unique apps targeted. |
| **CA Policies** | Per-policy detail: state, included/excluded users, groups, roles, applications, locations, platforms, grant controls, grant operator, timestamps. |
| **Named Locations** | All named locations including IP ranges (CIDR), country/region locations, and `IsTrusted` status. |
| **Application Scoping** | A flat row-per-(policy, application) view, with `Included` vs `Excluded` scope type. Disabled policies are skipped. |
| **App-Policy Mapping** | Per-application: how many enabled policies target it, and the names of those policies. Sorted by coverage. |

### Output files

The script writes to `-ExportPath` using filenames stamped with `yyyyMMdd-HHmmss`:

| Format | File(s) |
|--------|---------|
| `Excel` | `ConditionalAccess-Report-<timestamp>.xlsx` (5 worksheets) |
| `CSV` | `ConditionalAccess-Policies-<timestamp>.csv`, `ConditionalAccess-NamedLocations-<timestamp>.csv`, `ConditionalAccess-ApplicationScoping-<timestamp>.csv`, `ConditionalAccess-ApplicationMapping-<timestamp>.csv` |
| `JSON` | `ConditionalAccess-Report-<timestamp>.json` (single document with summary + all sections) |
| `HTML` | `ConditionalAccess-Report-<timestamp>.html` (interactive — tabbed, searchable, sortable, self-contained) |
| `All` | All of the above |

---

## Requirements

### Software

- Windows PowerShell 5.1 **or** PowerShell 7+
- Internet access to `graph.microsoft.com` and to the PowerShell Gallery (for first-run module install)

### PowerShell modules

The script will check for the following modules at startup and prompt to install any missing ones to the current user scope:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Identity.SignIns`
- `Microsoft.Graph.Applications`
- `Microsoft.Graph.Groups`
- `Microsoft.Graph.Users`
- `ImportExcel` (only strictly required for `Excel` / `All` formats, but always checked)

You can pre-install manually:

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Identity.SignIns, `
              Microsoft.Graph.Applications, Microsoft.Graph.Groups, Microsoft.Graph.Users, `
              ImportExcel -Scope CurrentUser
```

### Microsoft Graph permissions (delegated)

`Connect-MgGraph` is called with the following delegated scopes. The signed-in user must be able to consent to or already have these granted:

| Scope | Why it's needed |
|-------|-----------------|
| `Policy.Read.All` | Read Conditional Access policies and named locations. |
| `Application.Read.All` | Resolve Service Principal / Enterprise App display names from app IDs. |
| `Directory.Read.All` | Read directory role templates for role-targeted policies. |
| `Group.Read.All` | Resolve included/excluded group display names. |
| `User.Read.All` | Resolve included/excluded user display names and UPNs. |

### Required Entra ID role

Any of the following roles can run the report against a tenant:

- **Global Reader** (recommended — least privilege for read-only reporting)
- **Security Reader**
- **Conditional Access Administrator**
- **Security Administrator**
- **Global Administrator**

> Global Reader is the recommended role. The script never writes to the tenant.

If admin consent for the listed Graph scopes has not been granted for your account, a **Global Administrator** (or **Privileged Role Administrator**) must consent the first time the script is run in the tenant.

---

## Usage

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-ExportPath` | string | `.` | Directory to write report files to. Created if it does not exist. |
| `-ExportFormat` | string | `Excel` | One of `Excel`, `CSV`, `JSON`, `HTML`, `All`. |

### Examples

Default — Excel workbook into the current directory:

```powershell
.\Get-ConditionalAccessReport.ps1
```

Excel workbook into a specific folder:

```powershell
.\Get-ConditionalAccessReport.ps1 -ExportPath "C:\Reports" -ExportFormat Excel
```

Interactive HTML for sharing with a non-technical reviewer:

```powershell
.\Get-ConditionalAccessReport.ps1 -ExportFormat HTML -ExportPath "C:\Reports\CA"
```

Everything (Excel + CSV + JSON + HTML) — useful for evidence packages:

```powershell
.\Get-ConditionalAccessReport.ps1 -ExportFormat All -ExportPath "C:\Audit\2026-Q2"
```

CSV only — for ingestion into another reporting pipeline:

```powershell
.\Get-ConditionalAccessReport.ps1 -ExportFormat CSV
```

### Sign-in

On launch, the script triggers an interactive Microsoft Graph sign-in (device code or browser, depending on host). Sign in as an account that holds one of the supported roles in the target tenant.

---

## Interpreting the report

- **Application Scoping** intentionally excludes disabled policies — these have no effect at runtime.
- An app appearing in **App-Policy Mapping** with `Number of Policies = 0` will not show up; only apps explicitly included by an enabled policy are listed. Apps governed solely via `All` cloud apps will appear under the literal `All cloud apps` row.
- `"All users"`, `"All cloud apps"`, `"All locations"`, and `"All trusted locations"` are surfaced verbatim instead of being expanded into membership.
- Role assignments are resolved against **directory role templates**, so display names match the role catalog.

---

## Operational notes

- **Read-only.** The script never modifies policies, users, groups, or applications.
- **Tenant-scoped.** It reports on the tenant of the signed-in account. Connect with a different account to report on a different tenant.
- **No secrets are written to disk.** Output files contain policy and directory object metadata only.
- **Disconnects cleanly.** `Disconnect-MgGraph` is called at the end of every successful run.
- **Performance.** Runtime scales with the number of policies and the number of distinct users/groups/roles referenced. For large tenants, prefer the Excel or JSON formats — the HTML report inlines all rows into a single file.

---

## Troubleshooting

| Symptom | Likely cause | Resolution |
|---------|--------------|------------|
| `Failed to connect` at `Connect-MgGraph` | Missing scopes or no admin consent | Have a Global Admin consent the listed scopes, or sign in as a user with the consented permissions. |
| `Unknown App (<guid>)` rows | Service Principal not present in tenant (e.g., first-party app not yet provisioned) | Expected. The app ID is preserved so it can be cross-referenced manually. |
| `Group not found` / `User not found` | Object referenced by the policy was deleted | Expected. Update or clean up the policy. |
| Excel export fails | `ImportExcel` not installed | Re-run and accept the install prompt, or run `Install-Module ImportExcel -Scope CurrentUser`. |
| Module install fails with TLS error on older Windows | PowerShell defaulting to TLS 1.0 | Run `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12` before the script. |

---

## Security considerations

- Output files describe **who** is included/excluded from **which** policies — treat the artifacts as sensitive. Store them in an access-controlled location.
- Prefer **Global Reader** over higher-privilege roles when running this script.
- Review which accounts have standing consent for `Policy.Read.All` and the related scopes periodically.
