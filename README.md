# Microsoft-Tools

A collection of PowerShell scripts and tooling for administering Microsoft environments — Entra ID (Azure AD), Microsoft Graph, and Active Directory Domain Controllers.

Each tool lives in its own subfolder with a dedicated `README.md` that documents required Entra ID roles, Microsoft Graph scopes, PowerShell modules, parameters, usage examples, output, and troubleshooting.

---

## Tools

### Entra ID — [`ENTRA-TOOLS/`](./ENTRA-TOOLS/)

| Tool | Purpose | Docs |
|------|---------|------|
| [`ConditionalAccess-Report/`](./ENTRA-TOOLS/ConditionalAccess-Report/) | Inventory all Conditional Access policies, named locations, and the Enterprise Applications they target. Exports to Excel, CSV, JSON, or interactive HTML. | [README](./ENTRA-TOOLS/ConditionalAccess-Report/README.md) |
| [`FederatedTokenValidation-Assessment/`](./ENTRA-TOOLS/FederatedTokenValidation-Assessment/) | Read-only assessment of tenant exposure to Microsoft's stricter `federatedTokenValidationPolicy` enforcement. Identifies users likely to fail with `AADSTS5000820`. | [README](./ENTRA-TOOLS/FederatedTokenValidation-Assessment/README.md) |

### Domain Controllers — [`DC-TOOLS/`](./DC-TOOLS/)

| Tool | Purpose | Docs |
|------|---------|------|
| [`LDAPS-Renewal/`](./DC-TOOLS/LDAPS-Renewal/) | Automated LDAPS certificate lifecycle management for Domain Controllers using a Microsoft Enterprise CA. Includes installer, uninstaller, and renewal engine. | [README](./DC-TOOLS/LDAPS-Renewal/README.md) · [Admin Guide](./DC-TOOLS/LDAPS-Renewal/ADMIN-GUIDE.md) |

---

## Repository layout

```
Microsoft-Tools/
├── README.md                                       ← you are here
├── DC-TOOLS/
│   └── LDAPS-Renewal/
│       ├── Renew-LdapsCert.ps1
│       ├── Install-LdapsRenewTask.ps1
│       ├── Uninstall-LdapsRenewTask.ps1
│       ├── README.md
│       └── ADMIN-GUIDE.md
└── ENTRA-TOOLS/
    ├── ConditionalAccess-Report/
    │   ├── Get-ConditionalAccessReport.ps1
    │   └── README.md
    └── FederatedTokenValidation-Assessment/
        ├── Test-EntraFederatedTokenValidationImpact.ps1
        └── README.md
```

---

## General notes

- All scripts are PowerShell and intended to run on Windows (PowerShell 5.1 or PowerShell 7+).
- Entra scripts depend on the Microsoft Graph PowerShell SDK; most will offer to install required modules on first run.
- Review and test scripts in a non-production environment before deploying.
