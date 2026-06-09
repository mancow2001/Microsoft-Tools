# Test-EntraFederatedTokenValidationImpact.ps1

A **read-only** assessment script that identifies users in an Entra ID tenant who are likely to be impacted when Microsoft enforces stricter `federatedTokenValidationPolicy` checks on federated sign-ins.

This script does not change anything in the tenant. It collects domain, federation, and user data via Microsoft Graph, evaluates each user against a set of risk rules, and writes CSV reports to disk for review.

It is intended for Entra ID administrators of tenants that use **federated authentication** (AD FS, third-party IdPs via WS-Fed/SAML) and want to quantify exposure before Microsoft's enforcement window.

---

## Background

Microsoft is tightening validation of tokens issued by federated identity providers. Tenants are most at risk when:

- A user's **UPN domain** is *not* federated, but one of their **mail or proxy domains** *is* — and the IdP issues tokens against that alternate domain.
- A user's UPN domain is federated, but their mail domain (or alias users sign in with) is a different domain.
- The UPN domain is not a verified domain in the tenant.

These mismatches are silently tolerated today but can fail with `AADSTS5000820` (federated token validation) once enforcement is on.

---

## What it produces

All output is written to `-OutputDirectory` (default `.\EntraFederationAssessment`) with filenames stamped `yyyyMMdd-HHmmss`.

| File | Contents |
|------|----------|
| `domains-<ts>.csv` | All verified domains, authentication type (Federated / Managed), default/initial flags, supported services. |
| `federation-configurations-<ts>.csv` | Federation configuration per federated domain: IssuerUri, sign-in/sign-out URIs, preferred auth protocol, prompt behavior, presence of current and next signing certificates. |
| `user-federation-risk-assessment-<ts>.csv` | One row per user with computed `RiskLevel` (High / Medium / Informational / Low), `RiskReason`, UPN domain analysis, mail/proxy domain analysis, on-prem sync flags. |
| `summary-<ts>.csv` | Counts: domains, federated/managed domain lists, users assessed, risk totals. |
| `signin-federated-token-validation-failures-<ts>.csv` | *(Only when `-IncludeSignInLogs` is used.)* Recent sign-in failures matching `AADSTS5000820` or text indicating federated token validation issues. |

### Risk levels

| Level | Meaning |
|-------|---------|
| **High** | UPN domain is not a verified tenant domain **OR** the user has a federated mail/proxy domain that differs from their UPN domain **OR** the UPN domain is managed while an alternate login/email domain is federated. These users are the primary impact candidates. |
| **Medium** | UPN domain is federated, but the mail domain differs. Validate the sign-in identifier and IdP claim rules. |
| **Informational** | Alias/proxy domains differ from UPN. Usually normal — only investigate if users actually sign in with the alias. |
| **Low** | No mismatches detected. |

---

## Requirements

### Software

- Windows PowerShell 5.1 **or** PowerShell 7+
- Internet access to `graph.microsoft.com` and (on first run) the PowerShell Gallery

### PowerShell modules

Imported by the script:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Identity.DirectoryManagement`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Reports` *(only when `-IncludeSignInLogs` is used)*

Install once:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Or, if you prefer, let the script do it by passing `-InstallGraphModules`.

### Microsoft Graph permissions (delegated)

`Connect-MgGraph` is called with:

| Scope | Used for | Required? |
|-------|----------|-----------|
| `Directory.Read.All` | Read verified domains and federation configuration. | Always |
| `User.Read.All` | Read user UPN, mail, proxyAddresses, on-prem sync attributes. | Always |
| `AuditLog.Read.All` | Pull recent sign-in failures for AADSTS5000820. | Only when `-IncludeSignInLogs` is used |

> **Note on `Directory.Read.All` vs `Domain.Read.All`.** Reading the `/domains` and `/domains/{id}/federationConfiguration` Graph endpoints accepts **either** `Directory.Read.All` or `Domain.Read.All`. This script uses `Directory.Read.All` because it is almost always already admin-consented in enterprise tenants (Teams, Outlook, SharePoint, and most ISV apps depend on it), while `Domain.Read.All` typically requires an explicit, separate admin-consent step. If your tenant has `Domain.Read.All` consented and you prefer least-privilege, edit the `$scopes` array near the top of the script to swap them.

If admin consent for these scopes is not yet granted in the tenant, a Global Administrator (or Privileged Role Administrator) must consent on first run. To verify what is already consented: **Entra admin center** → **Enterprise applications** → **Microsoft Graph Command Line Tools** (App ID `14d82eec-204b-4c2f-b7e8-296a70dab67e`) → **Permissions**.

### Required Entra ID role

Any one of these will suffice:

- **Global Reader** (recommended — least privilege)
- **Security Reader**
- **Reports Reader** *(only if you need sign-in log access without other roles)*
- **Global Administrator**

> The script performs no writes. Choose the least-privileged role available.

### Licensing

- **Domain and user data**: included with all Entra ID tenants.
- **Sign-in logs (`-IncludeSignInLogs`)**: Entra ID **P1 or P2** is required to retrieve sign-in log data via Graph. Without P1/P2 the sign-in log call will fail; the script catches the failure, warns, and continues.

---

## Usage

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-OutputDirectory` | string | `.\EntraFederationAssessment` | Folder to write CSV reports to. Created if missing. |
| `-IncludeSignInLogs` | switch | off | Also pull recent sign-in failures and filter for AADSTS5000820 / federated-validation patterns. Requires `AuditLog.Read.All` and P1/P2. |
| `-DaysBack` | int | `30` | Lookback window for sign-in logs. Ignored unless `-IncludeSignInLogs` is set. |
| `-InstallGraphModules` | switch | off | Install the `Microsoft.Graph` meta-module to `CurrentUser` scope before importing. |

### Examples

Baseline assessment — domains, federation config, user risk:

```powershell
.\Test-EntraFederatedTokenValidationImpact.ps1
```

Custom output directory:

```powershell
.\Test-EntraFederatedTokenValidationImpact.ps1 -OutputDirectory "C:\Audit\Federation"
```

Include sign-in log evidence, last 30 days:

```powershell
.\Test-EntraFederatedTokenValidationImpact.ps1 -IncludeSignInLogs -DaysBack 30
```

Include sign-in log evidence, last 7 days, into a custom path:

```powershell
.\Test-EntraFederatedTokenValidationImpact.ps1 `
    -OutputDirectory "C:\Audit\Federation" `
    -IncludeSignInLogs `
    -DaysBack 7
```

First-time use on a fresh workstation (install Graph modules automatically):

```powershell
.\Test-EntraFederatedTokenValidationImpact.ps1 -InstallGraphModules
```

### Sign-in

The script triggers an interactive Microsoft Graph sign-in. The signed-in account's tenant determines which tenant is assessed.

---

## Interpreting the output

Open `user-federation-risk-assessment-<ts>.csv` and filter by `RiskLevel`:

1. **Start with `High`.** These are the users most likely to break when validation tightens. Review each one:
   - Is their UPN domain expected to be a federated domain? If yes, why isn't it federated in Entra?
   - Do they sign in with their UPN, mail address, or an alias? Compare against IdP claim rules.
   - Is the alternate federated domain on the same IdP as the UPN domain?
2. **Then review `Medium`.** UPN is federated, but the mail domain doesn't match. Confirm the IdP's `IssuerUri` matches what Entra has for the UPN domain (use `federation-configurations-<ts>.csv`).
3. **`Informational`** rows are typically aliases. Investigate only if you know users sign in with those aliases.
4. **`signin-federated-token-validation-failures-<ts>.csv`** (when available) is direct evidence — any row there is already failing federated validation under current behavior.

### Cross-referencing

- `domains-<ts>.csv` → list of verified domains and which are `Federated` vs `Managed`.
- `federation-configurations-<ts>.csv` → IdP wiring per federated domain. Look for missing signing certificates, mismatched `IssuerUri` values, or stale `NextSigningCertificate`.

---

## Operational notes

- **Read-only.** No writes, no consent escalation, no policy changes.
- **Idempotent.** Run as often as needed; each run is a fresh, timestamped snapshot.
- **Tenant-scoped.** Reports against the tenant of the signed-in account.
- **Sign-in log filtering is client-side.** The script pulls failed sign-ins for the window and filters locally because nested status/errorCode filters in Graph are inconsistent. For very large tenants, consider tightening `-DaysBack`.
- The Graph context (`Connect-MgGraph`) is *not* explicitly disconnected at the end of the script — the session persists for any follow-up cmdlets you run in the same shell. Run `Disconnect-MgGraph` manually when finished.

---

## Troubleshooting

| Symptom | Likely cause | Resolution |
|---------|--------------|------------|
| `Insufficient privileges` on `Get-MgDomain` or `Get-MgUser` | Account lacks the listed scopes / role | Use Global Reader or Security Reader; have a Global Admin consent the scopes. |
| `Unable to collect sign-in logs` warning | Missing `AuditLog.Read.All`, no P1/P2 license, or sign-in log retention exhausted | Grant scope and confirm P1/P2; reduce `-DaysBack` if the call times out. |
| `Get-MgDomainFederationConfiguration` throws per domain | Domain marked federated in Entra but federation object is missing or broken | The script records the error in `federation-configurations-<ts>.csv` and continues. Investigate the domain separately. |
| Empty `signin-federated-token-validation-failures-<ts>.csv` | No `AADSTS5000820` failures in window | Good news — but still review High-risk users; enforcement hasn't tripped yet. |
| Module import fails on first run | Microsoft Graph modules not installed | Re-run with `-InstallGraphModules`, or `Install-Module Microsoft.Graph -Scope CurrentUser`. |

---

## Security considerations

- The user risk CSV contains UPNs, display names, mail addresses, and proxy aliases. Treat output files as sensitive and store them in an access-controlled location.
- Prefer **Global Reader** for execution.
- The script never writes credentials, tokens, or signing certificates to disk. Federation configuration is captured as the *presence* (`Present` / `Missing`) of certificates, not the certificate material itself.

---

## Recommended workflow

1. Run a baseline (no `-IncludeSignInLogs`) and review domain inventory and federation configuration.
2. Run with `-IncludeSignInLogs -DaysBack 30` to capture current evidence.
3. Triage `High` risk users with the identity team and the owners of the federated IdP.
4. Re-run weekly until High counts trend to zero, then before any planned enforcement date.
