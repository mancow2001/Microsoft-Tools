<#
.SYNOPSIS
    Assess Entra ID tenant exposure to Microsoft federatedTokenValidationPolicy enforcement.

.DESCRIPTION
    Read-only assessment for tenants using federated authentication.
    Identifies users whose UPN domain differs from mail/proxy domains where one of those
    alternate domains is federated. These are likely candidates for impact when Entra
    enforces stricter federated token validation.

.NOTES
    Requires Microsoft Graph PowerShell SDK.

    Minimum useful scopes:
      Domain.Read.All
      User.Read.All

    Optional for sign-in log review:
      AuditLog.Read.All

    Example:
      .\Test-EntraFederatedTokenValidationImpact.ps1 -IncludeSignInLogs -DaysBack 30

.OUTPUTS
    CSV files in the selected output directory.
#>

param(
    [string]$OutputDirectory = ".\EntraFederationAssessment",

    [switch]$IncludeSignInLogs,

    [int]$DaysBack = 30,

    [switch]$InstallGraphModules
)

$ErrorActionPreference = "Stop"

function Write-Section {
    param([string]$Message)
    Write-Host ""
    Write-Host "==== $Message ====" -ForegroundColor Cyan
}

function Get-DomainSuffix {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    if ($Value -notmatch "@") {
        return $null
    }

    return ($Value.Split("@")[-1]).ToLowerInvariant()
}

function Normalize-ProxyAddress {
    param([string]$ProxyAddress)

    if ([string]::IsNullOrWhiteSpace($ProxyAddress)) {
        return $null
    }

    # Handles SMTP:user@domain.com, smtp:user@domain.com, SIP:user@domain.com, etc.
    if ($ProxyAddress -match "^[a-zA-Z]+:(.+@.+)$") {
        return $Matches[1]
    }

    if ($ProxyAddress -match ".+@.+") {
        return $ProxyAddress
    }

    return $null
}

function Safe-Join {
    param([object[]]$Items)

    if ($null -eq $Items -or $Items.Count -eq 0) {
        return ""
    }

    return (($Items | Where-Object { $_ } | Sort-Object -Unique) -join ";")
}

if ($InstallGraphModules) {
    Write-Section "Installing Microsoft Graph PowerShell modules"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Write-Section "Loading Microsoft Graph modules"
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

if ($IncludeSignInLogs) {
    Import-Module Microsoft.Graph.Reports -ErrorAction Stop
}

if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$scopes = @(
    "Domain.Read.All",
    "User.Read.All"
)

if ($IncludeSignInLogs) {
    $scopes += "AuditLog.Read.All"
}

Write-Section "Connecting to Microsoft Graph"
Connect-MgGraph -Scopes $scopes -NoWelcome

$context = Get-MgContext
Write-Host "Connected tenant: $($context.TenantId)"
Write-Host "Account:          $($context.Account)"
Write-Host "Scopes:           $($context.Scopes -join ', ')"

Write-Section "Collecting verified domains"
$domains = Get-MgDomain -All -Property Id,AuthenticationType,IsVerified,IsDefault,IsInitial,SupportedServices

$domainReport = foreach ($domain in $domains) {
    [pscustomobject]@{
        Domain              = $domain.Id.ToLowerInvariant()
        AuthenticationType  = $domain.AuthenticationType
        IsVerified          = $domain.IsVerified
        IsDefault           = $domain.IsDefault
        IsInitial           = $domain.IsInitial
        SupportedServices   = Safe-Join $domain.SupportedServices
    }
}

$domainReportPath = Join-Path $OutputDirectory "domains-$timestamp.csv"
$domainReport | Export-Csv -Path $domainReportPath -NoTypeInformation

$federatedDomains = @(
    $domainReport |
        Where-Object { $_.AuthenticationType -eq "Federated" -and $_.IsVerified -eq $true } |
        Select-Object -ExpandProperty Domain
)

$managedDomains = @(
    $domainReport |
        Where-Object { $_.AuthenticationType -eq "Managed" -and $_.IsVerified -eq $true } |
        Select-Object -ExpandProperty Domain
)

Write-Host "Verified domains:   $($domainReport.Count)"
Write-Host "Federated domains:  $($federatedDomains.Count)"
Write-Host "Managed domains:    $($managedDomains.Count)"

Write-Section "Collecting federation configuration objects"

$federationConfigResults = @()

foreach ($domain in $federatedDomains) {
    try {
        $configs = Get-MgDomainFederationConfiguration -DomainId $domain -All

        foreach ($config in $configs) {
            $federationConfigResults += [pscustomobject]@{
                Domain                         = $domain
                FederationConfigId             = $config.Id
                DisplayName                    = $config.DisplayName
                IssuerUri                      = $config.IssuerUri
                MetadataExchangeUri            = $config.MetadataExchangeUri
                PassiveSignInUri               = $config.PassiveSignInUri
                ActiveSignInUri                = $config.ActiveSignInUri
                SignOutUri                     = $config.SignOutUri
                PreferredAuthenticationProtocol = $config.PreferredAuthenticationProtocol
                PromptLoginBehavior            = $config.PromptLoginBehavior
                SigningCertificate             = if ($config.SigningCertificate) { "Present" } else { "Missing" }
                NextSigningCertificate         = if ($config.NextSigningCertificate) { "Present" } else { "Missing" }
            }
        }
    }
    catch {
        $federationConfigResults += [pscustomobject]@{
            Domain                         = $domain
            FederationConfigId             = ""
            DisplayName                    = ""
            IssuerUri                      = ""
            MetadataExchangeUri            = ""
            PassiveSignInUri               = ""
            ActiveSignInUri                = ""
            SignOutUri                     = ""
            PreferredAuthenticationProtocol = ""
            PromptLoginBehavior            = ""
            SigningCertificate             = ""
            NextSigningCertificate         = ""
            Error                          = $_.Exception.Message
        }
    }
}

$federationConfigPath = Join-Path $OutputDirectory "federation-configurations-$timestamp.csv"
$federationConfigResults | Export-Csv -Path $federationConfigPath -NoTypeInformation

Write-Host "Federation configurations collected: $($federationConfigResults.Count)"

Write-Section "Collecting users"

$userProperties = @(
    "id",
    "displayName",
    "userPrincipalName",
    "mail",
    "proxyAddresses",
    "accountEnabled",
    "userType",
    "onPremisesSyncEnabled",
    "onPremisesImmutableId"
)

$users = Get-MgUser -All -Property ($userProperties -join ",")

Write-Host "Users collected: $($users.Count)"

Write-Section "Evaluating UPN, mail, proxy, and federated-domain alignment"

$riskResults = foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    $upnDomain = Get-DomainSuffix $upn
    $mailDomain = Get-DomainSuffix $user.Mail

    $proxyEmailAddresses = @()
    $proxyDomains = @()

    if ($user.ProxyAddresses) {
        foreach ($proxy in $user.ProxyAddresses) {
            $normalizedProxy = Normalize-ProxyAddress $proxy

            if ($normalizedProxy) {
                $proxyEmailAddresses += $normalizedProxy.ToLowerInvariant()
                $proxyDomain = Get-DomainSuffix $normalizedProxy

                if ($proxyDomain) {
                    $proxyDomains += $proxyDomain
                }
            }
        }
    }

    $allAlternateDomains = @()
    if ($mailDomain) {
        $allAlternateDomains += $mailDomain
    }
    if ($proxyDomains) {
        $allAlternateDomains += $proxyDomains
    }

    $allAlternateDomains = @($allAlternateDomains | Where-Object { $_ } | Sort-Object -Unique)

    $federatedAlternateDomains = @(
        $allAlternateDomains |
            Where-Object { $federatedDomains -contains $_ }
    )

    $managedAlternateDomains = @(
        $allAlternateDomains |
            Where-Object { $managedDomains -contains $_ }
    )

    $isUpnDomainFederated = $federatedDomains -contains $upnDomain
    $isUpnDomainManaged = $managedDomains -contains $upnDomain
    $upnDomainIsVerified = ($domainReport.Domain -contains $upnDomain)

    $mailDiffersFromUpn = $false
    if ($mailDomain -and $upnDomain -and $mailDomain -ne $upnDomain) {
        $mailDiffersFromUpn = $true
    }

    $hasFederatedAlternateDomainDifferentFromUpn = $false
    foreach ($altDomain in $federatedAlternateDomains) {
        if ($altDomain -ne $upnDomain) {
            $hasFederatedAlternateDomainDifferentFromUpn = $true
        }
    }

    $proxyDomainsDifferentFromUpn = @(
        $proxyDomains |
            Where-Object { $_ -and $_ -ne $upnDomain } |
            Sort-Object -Unique
    )

    $riskLevel = "Low"
    $riskReason = @()

    if (-not $upnDomainIsVerified) {
        $riskLevel = "High"
        $riskReason += "UPN domain is not a verified Entra domain"
    }

    if ($hasFederatedAlternateDomainDifferentFromUpn) {
        $riskLevel = "High"
        $riskReason += "User has mail/proxy domain that is federated but differs from UPN domain"
    }

    if ($isUpnDomainManaged -and $hasFederatedAlternateDomainDifferentFromUpn) {
        $riskLevel = "High"
        $riskReason += "UPN domain is managed while alternate login/email domain is federated"
    }

    if ($isUpnDomainFederated -and $mailDiffersFromUpn) {
        if ($riskLevel -ne "High") {
            $riskLevel = "Medium"
        }
        $riskReason += "UPN domain is federated but mail domain differs; validate actual sign-in identifier and IdP claims"
    }

    if ($proxyDomainsDifferentFromUpn.Count -gt 0 -and $riskLevel -eq "Low") {
        $riskLevel = "Informational"
        $riskReason += "Proxy domains differ from UPN domain; likely normal for aliases, but validate if users sign in with aliases"
    }

    [pscustomobject]@{
        RiskLevel                                  = $riskLevel
        RiskReason                                 = Safe-Join $riskReason
        UserPrincipalName                          = $upn
        DisplayName                                = $user.DisplayName
        AccountEnabled                             = $user.AccountEnabled
        UserType                                   = $user.UserType
        UpnDomain                                  = $upnDomain
        UpnDomainVerified                          = $upnDomainIsVerified
        UpnDomainAuthenticationType                = if ($isUpnDomainFederated) { "Federated" } elseif ($isUpnDomainManaged) { "Managed" } else { "Unknown/Unverified" }
        Mail                                       = $user.Mail
        MailDomain                                 = $mailDomain
        MailDomainDiffersFromUpnDomain             = $mailDiffersFromUpn
        FederatedAlternateDomainsDifferentFromUpn  = Safe-Join (@($federatedAlternateDomains | Where-Object { $_ -ne $upnDomain }))
        ManagedAlternateDomains                    = Safe-Join $managedAlternateDomains
        ProxyDomainsDifferentFromUpn               = Safe-Join $proxyDomainsDifferentFromUpn
        ProxyAddresses                             = Safe-Join $proxyEmailAddresses
        OnPremisesSyncEnabled                      = $user.OnPremisesSyncEnabled
        OnPremisesImmutableIdPresent               = if ($user.OnPremisesImmutableId) { $true } else { $false }
        UserId                                     = $user.Id
    }
}

$riskReportPath = Join-Path $OutputDirectory "user-federation-risk-assessment-$timestamp.csv"
$riskResults | Export-Csv -Path $riskReportPath -NoTypeInformation

$highRisk = @($riskResults | Where-Object { $_.RiskLevel -eq "High" })
$mediumRisk = @($riskResults | Where-Object { $_.RiskLevel -eq "Medium" })
$infoRisk = @($riskResults | Where-Object { $_.RiskLevel -eq "Informational" })

Write-Host "High risk users:          $($highRisk.Count)" -ForegroundColor Red
Write-Host "Medium risk users:        $($mediumRisk.Count)" -ForegroundColor Yellow
Write-Host "Informational findings:   $($infoRisk.Count)" -ForegroundColor Gray

Write-Section "Creating summary report"

$summary = @(
    [pscustomobject]@{
        Category = "Tenant"
        Metric   = "TenantId"
        Value    = $context.TenantId
    },
    [pscustomobject]@{
        Category = "Domains"
        Metric   = "Verified domains"
        Value    = $domainReport.Count
    },
    [pscustomobject]@{
        Category = "Domains"
        Metric   = "Federated domains"
        Value    = $federatedDomains.Count
    },
    [pscustomobject]@{
        Category = "Domains"
        Metric   = "Managed domains"
        Value    = $managedDomains.Count
    },
    [pscustomobject]@{
        Category = "Users"
        Metric   = "Total users assessed"
        Value    = $users.Count
    },
    [pscustomobject]@{
        Category = "Risk"
        Metric   = "High risk users"
        Value    = $highRisk.Count
    },
    [pscustomobject]@{
        Category = "Risk"
        Metric   = "Medium risk users"
        Value    = $mediumRisk.Count
    },
    [pscustomobject]@{
        Category = "Risk"
        Metric   = "Informational users"
        Value    = $infoRisk.Count
    },
    [pscustomobject]@{
        Category = "FederatedDomains"
        Metric   = "Federated domain list"
        Value    = Safe-Join $federatedDomains
    },
    [pscustomobject]@{
        Category = "ManagedDomains"
        Metric   = "Managed domain list"
        Value    = Safe-Join $managedDomains
    }
)

$summaryPath = Join-Path $OutputDirectory "summary-$timestamp.csv"
$summary | Export-Csv -Path $summaryPath -NoTypeInformation

if ($IncludeSignInLogs) {
    Write-Section "Collecting recent sign-in failures for AADSTS5000820"

    $startDate = (Get-Date).AddDays(-1 * $DaysBack).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    try {
        # Pull recent failed sign-ins. Filtering nested status/errorCode server-side can be inconsistent,
        # so this pulls failures and filters locally.
        $signIns = Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $startDate"

        $federatedValidationFailures = @(
            $signIns |
                Where-Object {
                    $_.Status.ErrorCode -eq 5000820 -or
                    $_.Status.FailureReason -match "federated token validation|5000820|domain"
                } |
                Select-Object `
                    CreatedDateTime,
                    UserDisplayName,
                    UserPrincipalName,
                    AppDisplayName,
                    ClientAppUsed,
                    IpAddress,
                    ConditionalAccessStatus,
                    @{Name="ErrorCode";Expression={$_.Status.ErrorCode}},
                    @{Name="FailureReason";Expression={$_.Status.FailureReason}},
                    CorrelationId,
                    Id
        )

        $signInPath = Join-Path $OutputDirectory "signin-federated-token-validation-failures-$timestamp.csv"
        $federatedValidationFailures | Export-Csv -Path $signInPath -NoTypeInformation

        Write-Host "AADSTS5000820 / federated validation failures found: $($federatedValidationFailures.Count)" -ForegroundColor Yellow
    }
    catch {
        Write-Warning "Unable to collect sign-in logs. This may require Entra licensing, permissions, or AuditLog.Read.All admin consent."
        Write-Warning $_.Exception.Message
    }
}

Write-Section "Assessment complete"

Write-Host "Output directory:"
Write-Host "  $OutputDirectory"

Write-Host ""
Write-Host "Generated files:"
Write-Host "  $domainReportPath"
Write-Host "  $federationConfigPath"
Write-Host "  $riskReportPath"
Write-Host "  $summaryPath"

if ($IncludeSignInLogs) {
    Write-Host "  Sign-in log report generated if permissions/licensing allowed."
}

Write-Host ""
Write-Host "Interpretation:"
Write-Host "  High risk means the user has a federated mail/proxy/login domain that differs from their UPN domain."
Write-Host "  Medium risk means the user's UPN domain is federated, but the mail domain differs and should be validated."
Write-Host "  Informational means aliases/proxy domains differ from UPN; this may be normal unless users sign in with those aliases."
Write-Host ""
Write-Host "Recommended next step:"
Write-Host "  Validate high-risk users against actual IdP claim rules and sign-in behavior."