<#
.SYNOPSIS
    Manages LDAPS certificate lifecycle on Domain Controllers with automated renewal.

.DESCRIPTION
    Production-grade script for LDAPS certificate management on Domain Controllers.
    Supports three operational states:
    - State A: Valid cert exists, renew only if within threshold
    - State B: Expired/invalid cert exists, enroll immediately
    - State C: No LDAPS cert exists, bootstrap enrollment

    Designed to run as SYSTEM via Scheduled Task. No credentials stored.
    Idempotent and rollback-friendly with -WhatIf support.

    Supports automatic CA discovery from Active Directory when -CAConfig is not specified.

.PARAMETER CAConfig
    Optional. CA configuration string (e.g., "CAHOST\CA-NAME").
    If not specified, auto-discovers Enterprise CAs from Active Directory.

.PARAMETER TemplateName
    Certificate template name. Default: "LDAPS"

.PARAMETER BaseDomain
    Optional additional SAN DNS entry for base domain (e.g., "domain.com")

.PARAMETER IncludeShortNameSan
    Include DC hostname (short name) in SAN. Default: $true

.PARAMETER RenewWithinDays
    Days before expiration to trigger renewal. Default: 45

.PARAMETER LogPath
    Log file path. Default: C:\ProgramData\LdapsCertRenew\renew.log

.PARAMETER CleanupOld
    Remove superseded LDAPS certs after successful enrollment

.PARAMETER WhatIf
    Preview actions without making changes

.PARAMETER MinKeySize
    Minimum RSA key size. Default: 2048

.PARAMETER HashAlgorithm
    Hash algorithm for CSR. Default: sha256

.PARAMETER VerboseLogging
    Enable verbose/trace level logging for detailed troubleshooting

.PARAMETER PreferredCA
    When multiple CAs are discovered, prefer this CA name (partial match supported).
    Only used when -CAConfig is not specified.

.EXAMPLE
    .\Renew-LdapsCert.ps1 -BaseDomain "contoso.com"
    # Auto-discovers CA from Active Directory

.EXAMPLE
    .\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -BaseDomain "contoso.com"
    # Uses explicitly specified CA

.EXAMPLE
    .\Renew-LdapsCert.ps1 -PreferredCA "Issuing" -VerboseLogging
    # Auto-discovers CA, prefers one with "Issuing" in the name

.EXAMPLE
    .\Renew-LdapsCert.ps1 -CAConfig "CA01\Contoso-CA" -WhatIf -CleanupOld

.NOTES
    Version: 1.2.0
    Author: PKI Automation
    Requires: Windows Server 2016+, PowerShell 5.1+
    Run As: Local SYSTEM via Scheduled Task
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$CAConfig,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TemplateName = "LDAPS",

    [Parameter(Mandatory = $false)]
    [string]$PreferredCA,

    [Parameter(Mandatory = $false)]
    [string]$BaseDomain,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeShortNameSan = $true,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$RenewWithinDays = 45,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\ProgramData\LdapsCertRenew\renew.log",

    [Parameter(Mandatory = $false)]
    [switch]$CleanupOld,

    [Parameter(Mandatory = $false)]
    [ValidateSet(2048, 3072, 4096)]
    [int]$MinKeySize = 2048,

    [Parameter(Mandatory = $false)]
    [ValidateSet("sha256", "sha384", "sha512")]
    [string]$HashAlgorithm = "sha256",

    [Parameter(Mandatory = $false)]
    [switch]$VerboseLogging
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Constants
$script:SERVER_AUTH_OID = "1.3.6.1.5.5.7.3.1"
$script:SAN_OID = "2.5.29.17"
$script:EKU_OID = "2.5.29.37"
$script:WORK_DIR = "C:\ProgramData\LdapsCertRenew"
$script:ExitCode = 0
$script:ScriptVersion = "1.2.0"
$script:ScriptStartTime = Get-Date
$script:VerboseEnabled = $VerboseLogging.IsPresent -or $VerbosePreference -eq 'Continue'
#endregion

#region Logging Functions
function Initialize-Logging {
    [CmdletBinding()]
    param()

    $logDir = Split-Path -Path $LogPath -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }

    # Rotate log if > 10MB
    if (Test-Path -Path $LogPath) {
        $logFile = Get-Item -Path $LogPath
        if ($logFile.Length -gt 10MB) {
            $archivePath = $LogPath -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Move-Item -Path $LogPath -Destination $archivePath -Force
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG", "TRACE")]
        [string]$Level = "INFO"
    )

    # Skip TRACE level unless verbose logging is enabled
    if ($Level -eq "TRACE" -and -not $script:VerboseEnabled) {
        return
    }

    # Skip DEBUG level unless verbose logging is enabled (but always log DEBUG for key operations)
    if ($Level -eq "DEBUG" -and -not $script:VerboseEnabled) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
    $elapsedStr = "[+{0:F3}s]" -f $elapsed
    $logEntry = "[$timestamp] $elapsedStr [$Level] $Message"

    # Write to log file
    Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8

    # Also write to console for interactive debugging
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
        "DEBUG" { Write-Host $logEntry -ForegroundColor Gray }
        "TRACE" { Write-Host $logEntry -ForegroundColor DarkGray }
        default { Write-Host $logEntry }
    }
}

function Write-LogSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $separator = "=" * 70
    Write-Log -Message ""
    Write-Log -Message $separator
    Write-Log -Message $Title
    Write-Log -Message $separator
}

function Write-LogSubSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $separator = "-" * 50
    Write-Log -Message ""
    Write-Log -Message $separator
    Write-Log -Message $Title
    Write-Log -Message $separator
}

function Write-LogVerbose {
    <#
    .SYNOPSIS
        Writes verbose/trace level log entry for detailed troubleshooting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("DEBUG", "TRACE")]
        [string]$Level = "DEBUG"
    )

    Write-Log -Message $Message -Level $Level
}

function Write-LogObject {
    <#
    .SYNOPSIS
        Logs an object's properties for debugging.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [object]$Object,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties
    )

    if (-not $script:VerboseEnabled) { return }

    Write-LogVerbose -Message "Object dump: $Name" -Level TRACE

    if ($null -eq $Object) {
        Write-LogVerbose -Message "  (null)" -Level TRACE
        return
    }

    if ($Properties) {
        foreach ($prop in $Properties) {
            $value = $Object.$prop
            Write-LogVerbose -Message "  $prop = $value" -Level TRACE
        }
    }
    else {
        $Object.PSObject.Properties | ForEach-Object {
            Write-LogVerbose -Message "  $($_.Name) = $($_.Value)" -Level TRACE
        }
    }
}
#endregion

#region Environment Discovery Functions
function Get-EnvironmentInfo {
    <#
    .SYNOPSIS
        Gathers detailed environment information for troubleshooting.
    #>
    [CmdletBinding()]
    param()

    Write-LogSubSection -Title "Environment Information"

    # OS Information
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        Write-Log -Message "Operating System: $($os.Caption)"
        Write-Log -Message "OS Version: $($os.Version)"
        Write-Log -Message "OS Build: $($os.BuildNumber)"
        Write-LogVerbose -Message "OS Architecture: $($os.OSArchitecture)" -Level DEBUG
        Write-LogVerbose -Message "Last Boot: $($os.LastBootUpTime)" -Level DEBUG
    }
    catch {
        Write-Log -Message "Failed to get OS info: $_" -Level WARN
    }

    # PowerShell Version
    Write-Log -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-LogVerbose -Message "PowerShell Edition: $($PSVersionTable.PSEdition)" -Level DEBUG
    Write-LogVerbose -Message "CLR Version: $($PSVersionTable.CLRVersion)" -Level DEBUG

    # Current User Context
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    Write-Log -Message "Running as: $($currentIdentity.Name)"
    Write-LogVerbose -Message "User SID: $($currentIdentity.User.Value)" -Level DEBUG
    Write-LogVerbose -Message "Is System: $($currentIdentity.IsSystem)" -Level DEBUG
    Write-LogVerbose -Message "Auth Type: $($currentIdentity.AuthenticationType)" -Level DEBUG

    # Check if running elevated
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Log -Message "Running elevated: $isAdmin"

    # Domain Information
    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        Write-Log -Message "AD Domain: $($domain.Name)"
        Write-LogVerbose -Message "Forest: $($domain.Forest.Name)" -Level DEBUG
        Write-LogVerbose -Message "Domain Mode: $($domain.DomainMode)" -Level DEBUG
        Write-LogVerbose -Message "PDC: $($domain.PdcRoleOwner.Name)" -Level DEBUG
    }
    catch {
        Write-Log -Message "Failed to get domain info: $_" -Level WARN
    }

    # certreq.exe availability
    $certreqPath = Get-Command certreq.exe -ErrorAction SilentlyContinue
    if ($certreqPath) {
        Write-Log -Message "certreq.exe found: $($certreqPath.Source)"
        Write-LogVerbose -Message "certreq version info:" -Level DEBUG
        try {
            $versionInfo = (Get-Item $certreqPath.Source).VersionInfo
            Write-LogVerbose -Message "  File Version: $($versionInfo.FileVersion)" -Level DEBUG
            Write-LogVerbose -Message "  Product Version: $($versionInfo.ProductVersion)" -Level DEBUG
        }
        catch {
            Write-LogVerbose -Message "  Could not get version info" -Level DEBUG
        }
    }
    else {
        Write-Log -Message "certreq.exe NOT FOUND - enrollment will fail" -Level ERROR
    }

    # Certificate Store Access Test
    Write-LogVerbose -Message "Testing certificate store access..." -Level DEBUG
    try {
        $testCerts = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop
        Write-LogVerbose -Message "  LocalMachine\My accessible, contains $($testCerts.Count) certificates" -Level DEBUG
    }
    catch {
        Write-Log -Message "Cannot access LocalMachine\My certificate store: $_" -Level ERROR
    }

    # Network connectivity to CA (parse CA hostname from config)
    if (-not [string]::IsNullOrWhiteSpace($CAConfig) -and $CAConfig -match "^([^\\]+)\\") {
        $caHost = $Matches[1]
        Write-LogVerbose -Message "Testing connectivity to CA host: $caHost" -Level DEBUG
        try {
            $pingResult = Test-Connection -ComputerName $caHost -Count 1 -Quiet -ErrorAction SilentlyContinue
            if ($pingResult) {
                Write-Log -Message "CA host $caHost is reachable (ping)"
            }
            else {
                Write-Log -Message "CA host $caHost may not be reachable (ping failed)" -Level WARN
            }

            # Try RPC port
            $rpcTest = Test-NetConnection -ComputerName $caHost -Port 135 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            if ($rpcTest.TcpTestSucceeded) {
                Write-LogVerbose -Message "RPC port 135 open on $caHost" -Level DEBUG
            }
            else {
                Write-Log -Message "RPC port 135 not accessible on $caHost - CA enrollment may fail" -Level WARN
            }
        }
        catch {
            Write-Log -Message "Network test to CA failed: $_" -Level WARN
        }
    }
    elseif ([string]::IsNullOrWhiteSpace($CAConfig)) {
        Write-Log -Message "CA not specified - will auto-discover from Active Directory"
    }
}
#endregion

#region CA Auto-Discovery Functions
function Get-EnterpriseCAs {
    <#
    .SYNOPSIS
        Discovers Enterprise CAs from Active Directory.
    .DESCRIPTION
        Queries the Configuration partition for pKIEnrollmentService objects
        which represent Enterprise CAs in the forest.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TemplateName
    )

    Write-LogSubSection -Title "Enterprise CA Discovery"
    Write-Log -Message "Querying Active Directory for Enterprise CAs..."

    $enterpriseCAs = @()

    try {
        # Get the configuration naming context
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $configNC = $rootDSE.configurationNamingContext.Value
        Write-LogVerbose -Message "Configuration NC: $configNC" -Level DEBUG

        # Build the path to Enrollment Services
        $enrollmentServicesPath = "LDAP://CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNC"
        Write-LogVerbose -Message "Enrollment Services path: $enrollmentServicesPath" -Level DEBUG

        $enrollmentServices = [ADSI]$enrollmentServicesPath

        if ($null -eq $enrollmentServices -or $null -eq $enrollmentServices.Children) {
            Write-Log -Message "No Enrollment Services container found" -Level WARN
            return $enterpriseCAs
        }

        # Enumerate all pKIEnrollmentService objects (Enterprise CAs)
        foreach ($ca in $enrollmentServices.Children) {
            if ($ca.SchemaClassName -ne "pKIEnrollmentService") {
                continue
            }

            $caName = $ca.cn.Value
            $dnsHostname = $ca.dNSHostName.Value
            $caCertificateDN = $ca.cACertificateDN.Value
            $certificateTemplates = @($ca.certificateTemplates.Value)

            Write-LogVerbose -Message "Found CA: $caName" -Level DEBUG
            Write-LogVerbose -Message "  DNS Hostname: $dnsHostname" -Level DEBUG
            Write-LogVerbose -Message "  Certificate DN: $caCertificateDN" -Level DEBUG
            Write-LogVerbose -Message "  Templates published: $($certificateTemplates.Count)" -Level DEBUG

            # Build CA config string
            $caConfig = "$dnsHostname\$caName"

            # Check if the required template is published (if specified)
            $hasTemplate = $true
            if (-not [string]::IsNullOrWhiteSpace($TemplateName)) {
                $hasTemplate = $certificateTemplates -contains $TemplateName
                Write-LogVerbose -Message "  Has template '$TemplateName': $hasTemplate" -Level DEBUG
            }

            $caInfo = [PSCustomObject]@{
                Name              = $caName
                DNSHostName       = $dnsHostname
                Config            = $caConfig
                CertificateDN     = $caCertificateDN
                Templates         = $certificateTemplates
                HasRequiredTemplate = $hasTemplate
            }

            $enterpriseCAs += $caInfo

            Write-Log -Message "  Discovered CA: $caConfig"
            if (-not [string]::IsNullOrWhiteSpace($TemplateName)) {
                $templateStatus = if ($hasTemplate) { "YES" } else { "NO" }
                Write-Log -Message "    Template '$TemplateName' published: $templateStatus"
            }
        }

        Write-Log -Message "Total Enterprise CAs found: $($enterpriseCAs.Count)"

    }
    catch {
        Write-Log -Message "Failed to query Active Directory for CAs: $_" -Level ERROR
        Write-LogVerbose -Message "Exception type: $($_.Exception.GetType().FullName)" -Level DEBUG
        Write-LogVerbose -Message "Stack trace: $($_.ScriptStackTrace)" -Level DEBUG
    }

    return $enterpriseCAs
}

function Select-CertificateAuthority {
    <#
    .SYNOPSIS
        Selects the appropriate CA from discovered CAs.
    .DESCRIPTION
        Selection logic:
        1. If only one CA exists, use it
        2. If PreferredCA is specified, prefer matching CA
        3. If TemplateName specified, prefer CAs with that template
        4. Otherwise, use the first available CA
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$AvailableCAs,

        [Parameter(Mandatory = $false)]
        [string]$PreferredCA,

        [Parameter(Mandatory = $false)]
        [string]$TemplateName
    )

    Write-Log -Message "Selecting Certificate Authority..."

    if ($AvailableCAs.Count -eq 0) {
        Write-Log -Message "No Enterprise CAs available for selection" -Level ERROR
        return $null
    }

    if ($AvailableCAs.Count -eq 1) {
        $selected = $AvailableCAs[0]
        Write-Log -Message "Single CA available - auto-selected: $($selected.Config)"
        return $selected.Config
    }

    Write-Log -Message "Multiple CAs available ($($AvailableCAs.Count)) - applying selection criteria..."

    # Filter by template if specified
    $candidates = $AvailableCAs
    if (-not [string]::IsNullOrWhiteSpace($TemplateName)) {
        $withTemplate = $AvailableCAs | Where-Object { $_.HasRequiredTemplate }
        if ($withTemplate.Count -gt 0) {
            $candidates = @($withTemplate)
            Write-Log -Message "  Filtered to $($candidates.Count) CA(s) with template '$TemplateName'"
        }
        else {
            Write-Log -Message "  No CAs have template '$TemplateName' published - using all CAs" -Level WARN
        }
    }

    # Apply PreferredCA filter if specified
    if (-not [string]::IsNullOrWhiteSpace($PreferredCA)) {
        Write-Log -Message "  Applying preference filter: '$PreferredCA'"

        $preferred = $candidates | Where-Object {
            $_.Name -like "*$PreferredCA*" -or
            $_.DNSHostName -like "*$PreferredCA*" -or
            $_.Config -like "*$PreferredCA*"
        }

        if ($preferred.Count -gt 0) {
            $selected = $preferred[0]
            Write-Log -Message "  Preferred CA matched: $($selected.Config)"
            return $selected.Config
        }
        else {
            Write-Log -Message "  No CA matched preference '$PreferredCA' - using first available" -Level WARN
        }
    }

    # Default to first candidate
    $selected = $candidates[0]
    Write-Log -Message "  Selected CA: $($selected.Config)"
    return $selected.Config
}

function Resolve-CAConfiguration {
    <#
    .SYNOPSIS
        Resolves the CA configuration - either from parameter or auto-discovery.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExplicitCAConfig,

        [Parameter(Mandatory = $false)]
        [string]$PreferredCA,

        [Parameter(Mandatory = $false)]
        [string]$TemplateName
    )

    Write-LogSubSection -Title "CA Configuration Resolution"

    # If explicit CA provided, use it
    if (-not [string]::IsNullOrWhiteSpace($ExplicitCAConfig)) {
        Write-Log -Message "Using explicitly specified CA: $ExplicitCAConfig"
        return $ExplicitCAConfig
    }

    Write-Log -Message "No CA specified - initiating auto-discovery..."

    # Discover CAs from AD
    $discoveredCAs = Get-EnterpriseCAs -TemplateName $TemplateName

    if ($discoveredCAs.Count -eq 0) {
        Write-Log -Message "No Enterprise CAs discovered in Active Directory" -Level ERROR
        Write-Log -Message "Please specify -CAConfig parameter explicitly" -Level ERROR
        throw "CA auto-discovery failed: No Enterprise CAs found in Active Directory"
    }

    # Select appropriate CA
    $selectedCA = Select-CertificateAuthority -AvailableCAs $discoveredCAs `
        -PreferredCA $PreferredCA -TemplateName $TemplateName

    if ([string]::IsNullOrWhiteSpace($selectedCA)) {
        Write-Log -Message "Failed to select a Certificate Authority" -Level ERROR
        throw "CA auto-discovery failed: Could not select an appropriate CA"
    }

    Write-Log -Message ""
    Write-Log -Message "AUTO-DISCOVERED CA: $selectedCA"

    return $selectedCA
}
#endregion

#region Certificate Discovery Functions
function Get-DcIdentity {
    <#
    .SYNOPSIS
        Retrieves DC identity information for certificate operations.
    #>
    [CmdletBinding()]
    param()

    Write-LogSubSection -Title "DC Identity Discovery"

    Write-LogVerbose -Message "Querying Win32_ComputerSystem..." -Level TRACE
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem

    Write-LogObject -Name "Win32_ComputerSystem" -Object $computerSystem -Properties @("Name", "Domain", "DomainRole", "PartOfDomain")

    $dnsDomain = $computerSystem.Domain
    $hostname = $computerSystem.Name
    $fqdn = "$hostname.$dnsDomain"

    Write-Log -Message "DC Hostname: $hostname"
    Write-Log -Message "DNS Domain: $dnsDomain"
    Write-Log -Message "DC FQDN: $($fqdn.ToLower())"

    # Verify this is actually a DC
    $domainRole = $computerSystem.DomainRole
    Write-LogVerbose -Message "Domain Role: $domainRole (4=BDC, 5=PDC)" -Level DEBUG

    if ($domainRole -lt 4) {
        Write-Log -Message "WARNING: This machine may not be a Domain Controller (DomainRole=$domainRole)" -Level WARN
    }

    # Get additional DNS info
    Write-LogVerbose -Message "Querying DNS client configuration..." -Level TRACE
    try {
        $dnsConfig = Get-DnsClientGlobalSetting -ErrorAction SilentlyContinue
        if ($dnsConfig) {
            Write-LogVerbose -Message "DNS Suffix Search List: $($dnsConfig.SuffixSearchList -join ', ')" -Level DEBUG
        }
    }
    catch {
        Write-LogVerbose -Message "Could not query DNS config: $_" -Level DEBUG
    }

    return [PSCustomObject]@{
        Hostname   = $hostname
        DnsDomain  = $dnsDomain
        FQDN       = $fqdn.ToLower()
        DomainRole = $domainRole
    }
}

function Test-ServerAuthEku {
    <#
    .SYNOPSIS
        Tests if certificate has Server Authentication EKU.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    Write-LogVerbose -Message "  Checking EKU for cert: $($Certificate.Thumbprint)" -Level TRACE

    $ekuExtension = $Certificate.Extensions | Where-Object { $_.Oid.Value -eq $script:EKU_OID }
    if ($null -eq $ekuExtension) {
        Write-LogVerbose -Message "    No EKU extension found" -Level TRACE
        return $false
    }

    $eku = [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]$ekuExtension
    $ekuList = @()
    foreach ($oid in $eku.EnhancedKeyUsages) {
        $ekuList += "$($oid.FriendlyName) ($($oid.Value))"
        if ($oid.Value -eq $script:SERVER_AUTH_OID) {
            Write-LogVerbose -Message "    Found Server Authentication EKU" -Level TRACE
            return $true
        }
    }

    Write-LogVerbose -Message "    EKUs present: $($ekuList -join '; ')" -Level TRACE
    Write-LogVerbose -Message "    Server Authentication EKU NOT found" -Level TRACE
    return $false
}

function Get-CertificateSanEntries {
    <#
    .SYNOPSIS
        Extracts SAN DNS entries from certificate.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    Write-LogVerbose -Message "  Extracting SAN entries for cert: $($Certificate.Thumbprint)" -Level TRACE

    $sanEntries = @()
    $sanExtension = $Certificate.Extensions | Where-Object { $_.Oid.Value -eq $script:SAN_OID }

    if ($null -eq $sanExtension) {
        Write-LogVerbose -Message "    No SAN extension found" -Level TRACE
        return $sanEntries
    }

    # Get raw SAN string for debugging
    $sanStringMultiline = $sanExtension.Format($true)
    $sanStringSingleLine = $sanExtension.Format($false)

    Write-LogVerbose -Message "    Raw SAN (single line): $sanStringSingleLine" -Level TRACE

    # Extract DNS entries using regex
    $regexMatches = [regex]::Matches($sanStringSingleLine, 'DNS Name=([^\s,]+)', 'IgnoreCase')
    foreach ($match in $regexMatches) {
        $dnsName = $match.Groups[1].Value.ToLower()
        $sanEntries += $dnsName
        Write-LogVerbose -Message "    Extracted DNS SAN: $dnsName" -Level TRACE
    }

    if ($sanEntries.Count -eq 0) {
        Write-LogVerbose -Message "    No DNS SAN entries found in extension" -Level TRACE
    }

    return $sanEntries
}

function Get-CertificateDetails {
    <#
    .SYNOPSIS
        Gets comprehensive certificate details for logging.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    Write-LogVerbose -Message "  Certificate Details:" -Level DEBUG
    Write-LogVerbose -Message "    Thumbprint: $($Certificate.Thumbprint)" -Level DEBUG
    Write-LogVerbose -Message "    Subject: $($Certificate.Subject)" -Level DEBUG
    Write-LogVerbose -Message "    Issuer: $($Certificate.Issuer)" -Level DEBUG
    Write-LogVerbose -Message "    Serial: $($Certificate.SerialNumber)" -Level DEBUG
    Write-LogVerbose -Message "    NotBefore: $($Certificate.NotBefore)" -Level DEBUG
    Write-LogVerbose -Message "    NotAfter: $($Certificate.NotAfter)" -Level DEBUG
    Write-LogVerbose -Message "    HasPrivateKey: $($Certificate.HasPrivateKey)" -Level DEBUG
    Write-LogVerbose -Message "    SignatureAlgorithm: $($Certificate.SignatureAlgorithm.FriendlyName)" -Level DEBUG

    # Key info
    if ($Certificate.PublicKey) {
        Write-LogVerbose -Message "    PublicKey Algorithm: $($Certificate.PublicKey.Oid.FriendlyName)" -Level DEBUG
        Write-LogVerbose -Message "    PublicKey Size: $($Certificate.PublicKey.Key.KeySize) bits" -Level TRACE
    }

    # Extensions summary
    Write-LogVerbose -Message "    Extensions count: $($Certificate.Extensions.Count)" -Level TRACE
    foreach ($ext in $Certificate.Extensions) {
        Write-LogVerbose -Message "      Extension: $($ext.Oid.FriendlyName) ($($ext.Oid.Value)) Critical=$($ext.Critical)" -Level TRACE
    }

    # Private key info if available
    if ($Certificate.HasPrivateKey) {
        try {
            $privateKey = $Certificate.PrivateKey
            if ($privateKey) {
                Write-LogVerbose -Message "    PrivateKey Type: $($privateKey.GetType().Name)" -Level TRACE
                Write-LogVerbose -Message "    PrivateKey Exportable: Check via certutil" -Level TRACE
            }
        }
        catch {
            Write-LogVerbose -Message "    PrivateKey: Cannot access details" -Level TRACE
        }
    }
}

function Test-CertificateMatchesFqdn {
    <#
    .SYNOPSIS
        Tests if certificate Subject or SAN contains the DC FQDN.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [Parameter(Mandatory = $true)]
        [string]$FQDN
    )

    $fqdnLower = $FQDN.ToLower()
    Write-LogVerbose -Message "  Testing FQDN match for: $fqdnLower" -Level TRACE

    # Check Subject CN
    if ($Certificate.Subject -match "CN=([^,]+)") {
        $subjectCn = $Matches[1].Trim().ToLower()
        Write-LogVerbose -Message "    Subject CN: $subjectCn" -Level TRACE
        if ($subjectCn -eq $fqdnLower) {
            Write-LogVerbose -Message "    MATCH: Subject CN equals FQDN" -Level TRACE
            return $true
        }
    }
    else {
        Write-LogVerbose -Message "    No CN found in Subject: $($Certificate.Subject)" -Level TRACE
    }

    # Check SAN entries
    $sanEntries = Get-CertificateSanEntries -Certificate $Certificate
    Write-LogVerbose -Message "    SAN entries: $($sanEntries -join ', ')" -Level TRACE

    if ($sanEntries -contains $fqdnLower) {
        Write-LogVerbose -Message "    MATCH: FQDN found in SAN" -Level TRACE
        return $true
    }

    Write-LogVerbose -Message "    NO MATCH: FQDN not in Subject CN or SAN" -Level TRACE
    return $false
}

function Get-LdapsCandidateCertificates {
    <#
    .SYNOPSIS
        Discovers all LDAPS candidate certificates in LocalMachine\My.
    .DESCRIPTION
        A candidate certificate must:
        - Be in LocalMachine\My
        - Have a private key
        - Have Server Authentication EKU
        - Have Subject or SAN containing DC FQDN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DcFqdn
    )

    Write-LogSubSection -Title "Certificate Discovery"

    Write-Log -Message "Searching for LDAPS candidate certificates in LocalMachine\My..."
    Write-Log -Message "Target DC FQDN: $DcFqdn"
    Write-Log -Message "Candidate criteria:"
    Write-Log -Message "  - Has private key"
    Write-Log -Message "  - Has Server Authentication EKU (OID $script:SERVER_AUTH_OID)"
    Write-Log -Message "  - Subject CN or SAN contains: $DcFqdn"

    Write-LogVerbose -Message "Enumerating all certificates in LocalMachine\My..." -Level DEBUG

    $allCerts = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction SilentlyContinue

    if ($null -eq $allCerts -or $allCerts.Count -eq 0) {
        Write-Log -Message "No certificates found in LocalMachine\My store"
        return @()
    }

    Write-Log -Message "Total certificates in store: $($allCerts.Count)"
    $candidates = @()
    $certIndex = 0

    foreach ($cert in $allCerts) {
        $certIndex++
        Write-LogVerbose -Message "" -Level DEBUG
        Write-LogVerbose -Message "[$certIndex/$($allCerts.Count)] Evaluating certificate: $($cert.Thumbprint)" -Level DEBUG

        # Log full cert details in trace mode
        if ($script:VerboseEnabled) {
            Get-CertificateDetails -Certificate $cert
        }

        $isCandidate = $true
        $failReasons = @()
        $passReasons = @()

        # Check 1: Private key
        Write-LogVerbose -Message "  Check 1: Private key presence..." -Level TRACE
        if ($cert.HasPrivateKey) {
            $passReasons += "Has private key"
            Write-LogVerbose -Message "    PASS: Certificate has private key" -Level TRACE
        }
        else {
            $isCandidate = $false
            $failReasons += "No private key"
            Write-LogVerbose -Message "    FAIL: No private key" -Level TRACE
        }

        # Check 2: Server Authentication EKU
        Write-LogVerbose -Message "  Check 2: Server Authentication EKU..." -Level TRACE
        if (Test-ServerAuthEku -Certificate $cert) {
            $passReasons += "Has Server Auth EKU"
            Write-LogVerbose -Message "    PASS: Has Server Authentication EKU" -Level TRACE
        }
        else {
            $isCandidate = $false
            $failReasons += "Missing Server Auth EKU"
            Write-LogVerbose -Message "    FAIL: Missing Server Authentication EKU" -Level TRACE
        }

        # Check 3: FQDN match
        Write-LogVerbose -Message "  Check 3: FQDN match..." -Level TRACE
        if (Test-CertificateMatchesFqdn -Certificate $cert -FQDN $DcFqdn) {
            $passReasons += "FQDN matched"
            Write-LogVerbose -Message "    PASS: FQDN found in Subject/SAN" -Level TRACE
        }
        else {
            $isCandidate = $false
            $failReasons += "FQDN not in Subject/SAN"
            Write-LogVerbose -Message "    FAIL: FQDN not in Subject/SAN" -Level TRACE
        }

        # Summary for this cert
        if ($isCandidate) {
            $isExpired = $cert.NotAfter -lt (Get-Date)
            $daysRemaining = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
            $sanEntries = Get-CertificateSanEntries -Certificate $cert

            $candidateInfo = [PSCustomObject]@{
                Certificate   = $cert
                Thumbprint    = $cert.Thumbprint
                Subject       = $cert.Subject
                Issuer        = $cert.Issuer
                SerialNumber  = $cert.SerialNumber
                NotAfter      = $cert.NotAfter
                NotBefore     = $cert.NotBefore
                IsExpired     = $isExpired
                DaysRemaining = $daysRemaining
                SANEntries    = $sanEntries
            }
            $candidates += $candidateInfo

            $status = if ($isExpired) { "EXPIRED" } else { "Valid ($daysRemaining days remaining)" }
            Write-Log -Message "  CANDIDATE FOUND: $($cert.Thumbprint)"
            Write-Log -Message "    Subject: $($cert.Subject)"
            Write-Log -Message "    Issuer: $($cert.Issuer)"
            Write-Log -Message "    NotAfter: $($cert.NotAfter) - $status"
            Write-Log -Message "    SANs: $($sanEntries -join ', ')"
        }
        else {
            Write-LogVerbose -Message "  REJECTED: $($cert.Thumbprint)" -Level DEBUG
            Write-LogVerbose -Message "    Subject: $($cert.Subject)" -Level DEBUG
            Write-LogVerbose -Message "    Reasons: $($failReasons -join '; ')" -Level DEBUG
        }
    }

    Write-Log -Message ""
    Write-Log -Message "Discovery complete: $($candidates.Count) candidate(s) found out of $($allCerts.Count) total certificates"

    if ($candidates.Count -gt 0) {
        Write-Log -Message ""
        Write-Log -Message "Candidate summary:"
        $sortedCandidates = $candidates | Sort-Object -Property NotAfter -Descending
        $rank = 0
        foreach ($cand in $sortedCandidates) {
            $rank++
            $statusTag = if ($cand.IsExpired) { "[EXPIRED]" } else { "[VALID]" }
            Write-Log -Message "  #$rank $statusTag $($cand.Thumbprint) | NotAfter: $($cand.NotAfter) | Days: $($cand.DaysRemaining)"
        }
    }

    return $candidates
}

function Get-CertificateState {
    <#
    .SYNOPSIS
        Determines the certificate state and required action.
    .OUTPUTS
        PSCustomObject with State (A/B/C), BestCandidate, and ActionRequired
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Candidates,

        [Parameter(Mandatory = $true)]
        [int]$RenewThresholdDays,

        [Parameter(Mandatory = $false)]
        [string]$RequiredBaseDomain,

        [Parameter(Mandatory = $true)]
        [string]$DcFqdn,

        [Parameter(Mandatory = $true)]
        [bool]$RequireShortName,

        [Parameter(Mandatory = $true)]
        [string]$DcHostname
    )

    Write-LogSubSection -Title "State Determination"

    Write-Log -Message "Evaluating certificate state..."
    Write-Log -Message "  Candidate count: $($Candidates.Count)"
    Write-Log -Message "  Renewal threshold: $RenewThresholdDays days"
    Write-Log -Message "  Required DC FQDN in SAN: $DcFqdn"
    Write-Log -Message "  Require hostname in SAN: $RequireShortName (hostname: $DcHostname)"
    Write-Log -Message "  Required base domain in SAN: $(if ([string]::IsNullOrWhiteSpace($RequiredBaseDomain)) { '(none)' } else { $RequiredBaseDomain })"

    # State C: No candidates exist
    if ($Candidates.Count -eq 0) {
        Write-Log -Message ""
        Write-Log -Message ">>> STATE C: No LDAPS candidate certificates found"
        Write-Log -Message ">>> Action: Bootstrap enrollment required (first-time setup)"
        return [PSCustomObject]@{
            State           = "C"
            Description     = "No LDAPS certificate exists"
            BestCandidate   = $null
            ActionRequired  = $true
            Reason          = "First-time bootstrap enrollment - no valid LDAPS certificates in store"
        }
    }

    # Sort by NotAfter descending to get best candidate
    $sortedCandidates = $Candidates | Sort-Object -Property NotAfter -Descending
    $bestCandidate = $sortedCandidates[0]

    Write-Log -Message ""
    Write-Log -Message "Best candidate (by NotAfter):"
    Write-Log -Message "  Thumbprint: $($bestCandidate.Thumbprint)"
    Write-Log -Message "  Subject: $($bestCandidate.Subject)"
    Write-Log -Message "  NotBefore: $($bestCandidate.NotBefore)"
    Write-Log -Message "  NotAfter: $($bestCandidate.NotAfter)"
    Write-Log -Message "  IsExpired: $($bestCandidate.IsExpired)"
    Write-Log -Message "  DaysRemaining: $($bestCandidate.DaysRemaining)"
    Write-Log -Message "  SANEntries: $($bestCandidate.SANEntries -join ', ')"

    # Check if expired
    if ($bestCandidate.IsExpired) {
        Write-Log -Message ""
        Write-Log -Message ">>> STATE B: Best candidate certificate is EXPIRED"
        Write-Log -Message ">>> Expiration date: $($bestCandidate.NotAfter)"
        Write-Log -Message ">>> Action: Immediate enrollment required"
        return [PSCustomObject]@{
            State           = "B"
            Description     = "LDAPS certificate is expired"
            BestCandidate   = $bestCandidate
            ActionRequired  = $true
            Reason          = "Certificate expired on $($bestCandidate.NotAfter)"
        }
    }

    # Check SAN requirements
    Write-Log -Message ""
    Write-Log -Message "Checking SAN requirements..."

    $sanIssues = @()
    $fqdnLower = $DcFqdn.ToLower()
    $hostnameLower = $DcHostname.ToLower()

    # Check DC FQDN in SAN
    Write-LogVerbose -Message "  Checking for DC FQDN '$fqdnLower' in SAN..." -Level DEBUG
    if ($bestCandidate.SANEntries -contains $fqdnLower) {
        Write-Log -Message "  PASS: SAN contains DC FQDN: $fqdnLower"
    }
    else {
        $sanIssues += "Missing DC FQDN '$fqdnLower' in SAN"
        Write-Log -Message "  FAIL: SAN missing DC FQDN: $fqdnLower" -Level WARN
    }

    # Check hostname in SAN (if required)
    if ($RequireShortName) {
        Write-LogVerbose -Message "  Checking for hostname '$hostnameLower' in SAN..." -Level DEBUG
        if ($bestCandidate.SANEntries -contains $hostnameLower) {
            Write-Log -Message "  PASS: SAN contains hostname: $hostnameLower"
        }
        else {
            $sanIssues += "Missing hostname '$hostnameLower' in SAN"
            Write-Log -Message "  FAIL: SAN missing hostname: $hostnameLower" -Level WARN
        }
    }
    else {
        Write-LogVerbose -Message "  Hostname SAN check: SKIPPED (not required)" -Level DEBUG
    }

    # Check base domain in SAN (if specified)
    if (-not [string]::IsNullOrWhiteSpace($RequiredBaseDomain)) {
        $baseDomainLower = $RequiredBaseDomain.ToLower()
        Write-LogVerbose -Message "  Checking for base domain '$baseDomainLower' in SAN..." -Level DEBUG
        if ($bestCandidate.SANEntries -contains $baseDomainLower) {
            Write-Log -Message "  PASS: SAN contains base domain: $baseDomainLower"
        }
        else {
            $sanIssues += "Missing base domain '$baseDomainLower' in SAN"
            Write-Log -Message "  FAIL: SAN missing base domain: $baseDomainLower" -Level WARN
        }
    }
    else {
        Write-LogVerbose -Message "  Base domain SAN check: SKIPPED (not configured)" -Level DEBUG
    }

    if ($sanIssues.Count -gt 0) {
        Write-Log -Message ""
        Write-Log -Message ">>> STATE B: Certificate missing required SAN entries"
        Write-Log -Message ">>> Issues found:"
        foreach ($issue in $sanIssues) {
            Write-Log -Message ">>>   - $issue"
        }
        Write-Log -Message ">>> Action: Immediate enrollment required to get certificate with correct SANs"
        return [PSCustomObject]@{
            State           = "B"
            Description     = "LDAPS certificate missing required SAN entries"
            BestCandidate   = $bestCandidate
            ActionRequired  = $true
            Reason          = $sanIssues -join "; "
        }
    }

    # Check renewal threshold
    Write-Log -Message ""
    Write-Log -Message "Checking renewal threshold..."
    Write-Log -Message "  Days remaining: $($bestCandidate.DaysRemaining)"
    Write-Log -Message "  Threshold: $RenewThresholdDays days"

    if ($bestCandidate.DaysRemaining -le $RenewThresholdDays) {
        Write-Log -Message ""
        Write-Log -Message ">>> STATE A: Certificate within renewal threshold"
        Write-Log -Message ">>> Days remaining ($($bestCandidate.DaysRemaining)) <= Threshold ($RenewThresholdDays)"
        Write-Log -Message ">>> Action: Proactive renewal required"
        return [PSCustomObject]@{
            State           = "A"
            Description     = "LDAPS certificate within renewal threshold"
            BestCandidate   = $bestCandidate
            ActionRequired  = $true
            Reason          = "$($bestCandidate.DaysRemaining) days remaining (threshold: $RenewThresholdDays days)"
        }
    }

    # No action required
    Write-Log -Message ""
    Write-Log -Message ">>> STATE A: Certificate is valid and not within renewal threshold"
    Write-Log -Message ">>> Days remaining ($($bestCandidate.DaysRemaining)) > Threshold ($RenewThresholdDays)"
    Write-Log -Message ">>> Action: None required"
    return [PSCustomObject]@{
        State           = "A"
        Description     = "LDAPS certificate is valid"
        BestCandidate   = $bestCandidate
        ActionRequired  = $false
        Reason          = "$($bestCandidate.DaysRemaining) days remaining exceeds threshold of $RenewThresholdDays days"
    }
}
#endregion

#region Certificate Enrollment Functions
function New-CertificateRequestInf {
    <#
    .SYNOPSIS
        Generates INF file for certreq.exe.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubjectCN,

        [Parameter(Mandatory = $true)]
        [string[]]$SANDnsNames,

        [Parameter(Mandatory = $true)]
        [string]$TemplateName,

        [Parameter(Mandatory = $true)]
        [int]$KeySize,

        [Parameter(Mandatory = $true)]
        [string]$HashAlgorithm,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Log -Message "Generating certificate request INF file..."

    # Build SAN string for INF
    $sanString = ($SANDnsNames | ForEach-Object { "dns=$_" }) -join "&"

    $infContent = @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject = "CN=$SubjectCN"
KeySpec = 1
KeyLength = $KeySize
Exportable = FALSE
MachineKeySet = TRUE
SMIME = FALSE
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft Software Key Storage Provider"
ProviderType = 12
RequestType = PKCS10
KeyUsage = 0xa0
HashAlgorithm = $HashAlgorithm

[EnhancedKeyUsageExtension]
OID = $script:SERVER_AUTH_OID

[Extensions]
$script:SAN_OID = "{text}$sanString"

[RequestAttributes]
CertificateTemplate = $TemplateName
"@

    Set-Content -Path $OutputPath -Value $infContent -Encoding ASCII -Force

    # Log INF details
    Write-Log -Message "INF file generated: $OutputPath"
    Write-Log -Message "INF Configuration:"
    Write-Log -Message "  Subject: CN=$SubjectCN"
    Write-Log -Message "  SAN DNS entries: $($SANDnsNames -join ', ')"
    Write-Log -Message "  Template: $TemplateName"
    Write-Log -Message "  Key size: $KeySize bits"
    Write-Log -Message "  Hash algorithm: $HashAlgorithm"
    Write-Log -Message "  Key Storage Provider: Microsoft Software Key Storage Provider"
    Write-Log -Message "  Exportable: FALSE"
    Write-Log -Message "  MachineKeySet: TRUE"

    # Log full INF content in verbose mode
    Write-LogVerbose -Message "" -Level DEBUG
    Write-LogVerbose -Message "Full INF file contents:" -Level DEBUG
    Write-LogVerbose -Message "--- BEGIN INF ---" -Level DEBUG
    foreach ($line in ($infContent -split "`n")) {
        Write-LogVerbose -Message "  $line" -Level DEBUG
    }
    Write-LogVerbose -Message "--- END INF ---" -Level DEBUG

    return $infContent
}

function Invoke-CertReq {
    <#
    .SYNOPSIS
        Executes certreq.exe and captures output with detailed logging.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    Write-Log -Message ""
    Write-Log -Message "Executing certreq.exe operation: $Operation"
    Write-Log -Message "Command: certreq.exe $Arguments"
    Write-LogVerbose -Message "Working directory: $(Get-Location)" -Level DEBUG

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "certreq.exe"
    $processInfo.Arguments = $Arguments
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true
    $processInfo.WorkingDirectory = $script:WORK_DIR

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo

    Write-LogVerbose -Message "Starting process..." -Level TRACE
    $process.Start() | Out-Null

    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    $stopwatch.Stop()

    $result = [PSCustomObject]@{
        ExitCode     = $process.ExitCode
        StdOut       = $stdout
        StdErr       = $stderr
        Success      = ($process.ExitCode -eq 0)
        DurationMs   = $stopwatch.ElapsedMilliseconds
    }

    Write-Log -Message "certreq $Operation completed in $($result.DurationMs)ms"
    Write-Log -Message "Exit code: $($result.ExitCode) $(if ($result.Success) { '(SUCCESS)' } else { '(FAILED)' })"

    # Log stdout
    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        Write-Log -Message "Standard output:"
        foreach ($line in ($stdout -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $trimmedLine = $line.Trim()
                # Always log certreq output as it's important for troubleshooting
                Write-Log -Message "  [stdout] $trimmedLine"
            }
        }
    }
    else {
        Write-LogVerbose -Message "Standard output: (empty)" -Level DEBUG
    }

    # Log stderr
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        Write-Log -Message "Standard error:" -Level WARN
        foreach ($line in ($stderr -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Log -Message "  [stderr] $($line.Trim())" -Level WARN
            }
        }
    }
    else {
        Write-LogVerbose -Message "Standard error: (empty)" -Level DEBUG
    }

    # Parse and log specific information
    if ($stdout -match "RequestId:\s*(\d+)") {
        Write-Log -Message "Parsed CA Request ID: $($Matches[1])"
    }

    if ($stdout -match "Certificate retrieved") {
        Write-Log -Message "Certificate was successfully retrieved from CA"
    }

    if ($stdout -match "Installed Certificate") {
        Write-Log -Message "Certificate was installed into certificate store"
    }

    # Log error codes
    if (-not $result.Success) {
        Write-Log -Message "certreq failed with exit code $($result.ExitCode)" -Level ERROR

        # Common error codes
        switch ($result.ExitCode) {
            -2146762487 { Write-Log -Message "Error: The certificate template is not valid" -Level ERROR }
            -2146762486 { Write-Log -Message "Error: The requested certificate template is not supported" -Level ERROR }
            -2146877432 { Write-Log -Message "Error: Certificate request denied" -Level ERROR }
            -2146762480 { Write-Log -Message "Error: A required certificate is not within its validity period" -Level ERROR }
            5           { Write-Log -Message "Error: Certificate request is pending manager approval" -Level WARN }
            default     { Write-Log -Message "See Microsoft documentation for error code $($result.ExitCode)" -Level ERROR }
        }
    }

    return $result
}

function Request-LdapsCertificate {
    <#
    .SYNOPSIS
        Performs full certificate enrollment: generate CSR, submit to CA, accept certificate.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CAConfig,

        [Parameter(Mandatory = $true)]
        [string]$TemplateName,

        [Parameter(Mandatory = $true)]
        [string]$DcFqdn,

        [Parameter(Mandatory = $false)]
        [string]$DcHostname,

        [Parameter(Mandatory = $false)]
        [string]$BaseDomain,

        [Parameter(Mandatory = $true)]
        [bool]$IncludeShortName,

        [Parameter(Mandatory = $true)]
        [int]$KeySize,

        [Parameter(Mandatory = $true)]
        [string]$HashAlgorithm
    )

    Write-LogSection -Title "Certificate Enrollment"

    Write-Log -Message "Starting certificate enrollment process..."
    Write-Log -Message "CA Configuration: $CAConfig"
    Write-Log -Message "Template Name: $TemplateName"

    # Build SAN list
    Write-Log -Message ""
    Write-Log -Message "Building SAN (Subject Alternative Name) list..."

    $sanList = @($DcFqdn.ToLower())
    Write-Log -Message "  Added DC FQDN: $($DcFqdn.ToLower())"

    if ($IncludeShortName -and -not [string]::IsNullOrWhiteSpace($DcHostname)) {
        $hostnameEntry = $DcHostname.ToLower()
        if ($sanList -notcontains $hostnameEntry) {
            $sanList += $hostnameEntry
            Write-Log -Message "  Added hostname: $hostnameEntry"
        }
        else {
            Write-LogVerbose -Message "  Hostname already in list (duplicate of FQDN?)" -Level DEBUG
        }
    }
    else {
        Write-Log -Message "  Hostname: SKIPPED (IncludeShortName=$IncludeShortName)"
    }

    if (-not [string]::IsNullOrWhiteSpace($BaseDomain)) {
        $baseDomainEntry = $BaseDomain.ToLower()
        if ($sanList -notcontains $baseDomainEntry) {
            $sanList += $baseDomainEntry
            Write-Log -Message "  Added base domain: $baseDomainEntry"
        }
        else {
            Write-LogVerbose -Message "  Base domain already in list" -Level DEBUG
        }
    }
    else {
        Write-Log -Message "  Base domain: SKIPPED (not specified)"
    }

    Write-Log -Message ""
    Write-Log -Message "Final SAN list ($($sanList.Count) entries):"
    foreach ($san in $sanList) {
        Write-Log -Message "  - $san"
    }

    # Prepare file paths
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $infPath = Join-Path -Path $script:WORK_DIR -ChildPath "ldaps_request_$timestamp.inf"
    $reqPath = Join-Path -Path $script:WORK_DIR -ChildPath "ldaps_request_$timestamp.req"
    $cerPath = Join-Path -Path $script:WORK_DIR -ChildPath "ldaps_cert_$timestamp.cer"

    Write-Log -Message ""
    Write-Log -Message "Working files:"
    Write-Log -Message "  INF path: $infPath"
    Write-Log -Message "  REQ path: $reqPath"
    Write-Log -Message "  CER path: $cerPath"

    # Ensure work directory exists
    if (-not (Test-Path -Path $script:WORK_DIR)) {
        Write-Log -Message "Creating work directory: $script:WORK_DIR"
        New-Item -Path $script:WORK_DIR -ItemType Directory -Force | Out-Null
    }

    try {
        # Generate INF file
        $null = New-CertificateRequestInf -SubjectCN $DcFqdn -SANDnsNames $sanList `
            -TemplateName $TemplateName -KeySize $KeySize -HashAlgorithm $HashAlgorithm `
            -OutputPath $infPath

        if ($PSCmdlet.ShouldProcess("Generate PKCS#10 request from $infPath", "certreq -new")) {
            # Step 1: Generate CSR
            Write-LogSubSection -Title "Step 1/3: Generate Certificate Request (CSR)"
            Write-Log -Message "Creating PKCS#10 certificate signing request..."

            $newResult = Invoke-CertReq -Arguments "-new `"$infPath`" `"$reqPath`"" -Operation "new"

            if (-not $newResult.Success) {
                Write-Log -Message "CSR generation failed!" -Level ERROR
                throw "certreq -new failed with exit code $($newResult.ExitCode)"
            }

            if (-not (Test-Path -Path $reqPath)) {
                Write-Log -Message "CSR file was not created!" -Level ERROR
                throw "CSR file was not created: $reqPath"
            }

            $reqSize = (Get-Item -Path $reqPath).Length
            Write-Log -Message "CSR generated successfully"
            Write-Log -Message "  File: $reqPath"
            Write-Log -Message "  Size: $reqSize bytes"

            # Log CSR details in verbose mode
            if ($script:VerboseEnabled) {
                Write-LogVerbose -Message "CSR dump (certutil -dump):" -Level DEBUG
                try {
                    $dumpOutput = & certutil -dump "$reqPath" 2>&1
                    foreach ($line in ($dumpOutput -split "`n")) {
                        Write-LogVerbose -Message "  $line" -Level TRACE
                    }
                }
                catch {
                    Write-LogVerbose -Message "  Could not dump CSR: $_" -Level DEBUG
                }
            }
        }
        else {
            Write-Log -Message "[WhatIf] Would generate CSR: $reqPath"
            return [PSCustomObject]@{
                Success    = $true
                WhatIf     = $true
                Thumbprint = $null
                Message    = "WhatIf: Certificate request would be generated"
            }
        }

        if ($PSCmdlet.ShouldProcess("Submit CSR to CA: $CAConfig", "certreq -submit")) {
            # Step 2: Submit to CA
            Write-LogSubSection -Title "Step 2/3: Submit Request to CA"
            Write-Log -Message "Submitting certificate request to CA..."
            Write-Log -Message "CA: $CAConfig"

            $submitResult = Invoke-CertReq -Arguments "-submit -config `"$CAConfig`" `"$reqPath`" `"$cerPath`"" -Operation "submit"

            # Check for request ID
            $requestId = $null
            if ($submitResult.StdOut -match "RequestId:\s*(\d+)") {
                $requestId = $Matches[1]
                Write-Log -Message "CA assigned Request ID: $requestId"
            }

            if (-not $submitResult.Success) {
                # Check if it's pending
                if ($submitResult.ExitCode -eq 5 -or $submitResult.StdOut -match "pending") {
                    $pendingMsg = "Certificate request is pending CA manager approval"
                    Write-Log -Message $pendingMsg -Level WARN
                    Write-Log -Message "Request ID: $requestId" -Level WARN
                    Write-Log -Message "Action required: Approve the request in CA console, then re-run this script" -Level WARN
                    return [PSCustomObject]@{
                        Success    = $false
                        Pending    = $true
                        RequestId  = $requestId
                        Thumbprint = $null
                        Message    = $pendingMsg
                    }
                }
                Write-Log -Message "Certificate request submission failed!" -Level ERROR
                throw "certreq -submit failed with exit code $($submitResult.ExitCode)"
            }

            if (-not (Test-Path -Path $cerPath)) {
                Write-Log -Message "Certificate file was not created by CA!" -Level ERROR
                throw "Certificate file was not created: $cerPath"
            }

            $cerSize = (Get-Item -Path $cerPath).Length
            Write-Log -Message "Certificate issued successfully by CA"
            Write-Log -Message "  File: $cerPath"
            Write-Log -Message "  Size: $cerSize bytes"

            # Log certificate details in verbose mode
            if ($script:VerboseEnabled) {
                Write-LogVerbose -Message "Issued certificate dump (certutil -dump):" -Level DEBUG
                try {
                    $dumpOutput = & certutil -dump "$cerPath" 2>&1
                    foreach ($line in ($dumpOutput -split "`n")) {
                        Write-LogVerbose -Message "  $line" -Level TRACE
                    }
                }
                catch {
                    Write-LogVerbose -Message "  Could not dump certificate: $_" -Level DEBUG
                }
            }
        }
        else {
            Write-Log -Message "[WhatIf] Would submit CSR to CA: $CAConfig"
            return [PSCustomObject]@{
                Success    = $true
                WhatIf     = $true
                Thumbprint = $null
                Message    = "WhatIf: Certificate request would be submitted"
            }
        }

        if ($PSCmdlet.ShouldProcess("Accept certificate into LocalMachine\My", "certreq -accept")) {
            # Step 3: Accept/Install certificate
            Write-LogSubSection -Title "Step 3/3: Install Certificate"
            Write-Log -Message "Installing issued certificate into LocalMachine\My store..."

            $acceptResult = Invoke-CertReq -Arguments "-accept `"$cerPath`"" -Operation "accept"

            if (-not $acceptResult.Success) {
                Write-Log -Message "Certificate installation failed!" -Level ERROR
                throw "certreq -accept failed with exit code $($acceptResult.ExitCode)"
            }

            Write-Log -Message "Certificate installed successfully into LocalMachine\My"

            # Extract thumbprint from output or find newly installed cert
            $newThumbprint = $null
            if ($acceptResult.StdOut -match "Thumbprint:\s*([A-Fa-f0-9]{40})") {
                $newThumbprint = $Matches[1].ToUpper()
                Write-Log -Message "Thumbprint from certreq output: $newThumbprint"
            }
            else {
                Write-Log -Message "Thumbprint not in certreq output, searching certificate store..."
                # Wait briefly for store update
                Start-Sleep -Seconds 2
                $newCandidates = Get-LdapsCandidateCertificates -DcFqdn $DcFqdn
                if ($newCandidates.Count -gt 0) {
                    $newestCert = $newCandidates | Sort-Object -Property NotAfter -Descending | Select-Object -First 1
                    $newThumbprint = $newestCert.Thumbprint
                    Write-Log -Message "Found newest certificate: $newThumbprint"
                }
            }

            if ([string]::IsNullOrWhiteSpace($newThumbprint)) {
                Write-Log -Message "Unable to determine thumbprint of newly installed certificate!" -Level ERROR
                throw "Unable to determine thumbprint of newly installed certificate"
            }

            Write-Log -Message ""
            Write-Log -Message "Certificate enrollment completed successfully"
            Write-Log -Message "New certificate thumbprint: $newThumbprint"

            return [PSCustomObject]@{
                Success    = $true
                WhatIf     = $false
                Thumbprint = $newThumbprint
                Message    = "Certificate enrolled and installed successfully"
            }
        }
        else {
            Write-Log -Message "[WhatIf] Would accept certificate into store"
            return [PSCustomObject]@{
                Success    = $true
                WhatIf     = $true
                Thumbprint = $null
                Message    = "WhatIf: Certificate would be installed"
            }
        }
    }
    finally {
        # Log working file status
        Write-Log -Message ""
        Write-Log -Message "Working files status (retained for audit/troubleshooting):"
        if (Test-Path -Path $infPath) {
            Write-Log -Message "  INF: $infPath ($(( Get-Item $infPath).Length) bytes)"
        }
        if (Test-Path -Path $reqPath) {
            Write-Log -Message "  REQ: $reqPath ($((Get-Item $reqPath).Length) bytes)"
        }
        if (Test-Path -Path $cerPath) {
            Write-Log -Message "  CER: $cerPath ($((Get-Item $cerPath).Length) bytes)"
        }
    }
}
#endregion

#region Verification Functions
function Test-NewCertificateCompliance {
    <#
    .SYNOPSIS
        Verifies the newly installed certificate meets all requirements.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint,

        [Parameter(Mandatory = $true)]
        [string]$DcFqdn,

        [Parameter(Mandatory = $false)]
        [string]$BaseDomain,

        [Parameter(Mandatory = $true)]
        [bool]$RequireShortName,

        [Parameter(Mandatory = $true)]
        [string]$DcHostname,

        [Parameter(Mandatory = $false)]
        [datetime]$PreviousNotAfter
    )

    Write-LogSection -Title "Post-Installation Verification"

    Write-Log -Message "Verifying newly installed certificate..."
    Write-Log -Message "Target thumbprint: $Thumbprint"

    $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction SilentlyContinue
    if ($null -eq $cert) {
        Write-Log -Message "CRITICAL: Certificate not found in store!" -Level ERROR
        Write-Log -Message "Expected path: Cert:\LocalMachine\My\$Thumbprint" -Level ERROR
        return $false
    }

    Write-Log -Message "Certificate found in store"

    # Log full certificate details
    if ($script:VerboseEnabled) {
        Get-CertificateDetails -Certificate $cert
    }

    $allPassed = $true
    $checkNumber = 0

    # Check 1: Has private key
    $checkNumber++
    Write-Log -Message ""
    Write-Log -Message "Check $checkNumber`: Private key presence"
    if ($cert.HasPrivateKey) {
        Write-Log -Message "  [PASS] Certificate has private key"
    }
    else {
        Write-Log -Message "  [FAIL] Certificate does NOT have private key" -Level ERROR
        $allPassed = $false
    }

    # Check 2: Has Server Authentication EKU
    $checkNumber++
    Write-Log -Message ""
    Write-Log -Message "Check $checkNumber`: Server Authentication EKU"
    if (Test-ServerAuthEku -Certificate $cert) {
        Write-Log -Message "  [PASS] Certificate has Server Authentication EKU (OID $script:SERVER_AUTH_OID)"
    }
    else {
        Write-Log -Message "  [FAIL] Certificate missing Server Authentication EKU" -Level ERROR
        $allPassed = $false
    }

    # Check 3: SAN entries
    $checkNumber++
    Write-Log -Message ""
    Write-Log -Message "Check $checkNumber`: Subject Alternative Names (SAN)"
    $sanEntries = Get-CertificateSanEntries -Certificate $cert
    Write-Log -Message "  Found $($sanEntries.Count) SAN DNS entries:"
    foreach ($san in $sanEntries) {
        Write-Log -Message "    - $san"
    }

    # Check DC FQDN
    Write-Log -Message ""
    Write-Log -Message "  Checking for DC FQDN: $($DcFqdn.ToLower())"
    if ($sanEntries -contains $DcFqdn.ToLower()) {
        Write-Log -Message "    [PASS] DC FQDN found in SAN"
    }
    else {
        Write-Log -Message "    [FAIL] DC FQDN NOT found in SAN" -Level ERROR
        $allPassed = $false
    }

    # Check hostname if required
    if ($RequireShortName) {
        Write-Log -Message ""
        Write-Log -Message "  Checking for hostname: $($DcHostname.ToLower())"
        if ($sanEntries -contains $DcHostname.ToLower()) {
            Write-Log -Message "    [PASS] Hostname found in SAN"
        }
        else {
            Write-Log -Message "    [FAIL] Hostname NOT found in SAN" -Level ERROR
            $allPassed = $false
        }
    }

    # Check base domain if specified
    if (-not [string]::IsNullOrWhiteSpace($BaseDomain)) {
        Write-Log -Message ""
        Write-Log -Message "  Checking for base domain: $($BaseDomain.ToLower())"
        if ($sanEntries -contains $BaseDomain.ToLower()) {
            Write-Log -Message "    [PASS] Base domain found in SAN"
        }
        else {
            Write-Log -Message "    [FAIL] Base domain NOT found in SAN" -Level ERROR
            $allPassed = $false
        }
    }

    # Check 4: NotAfter is in the future
    $checkNumber++
    Write-Log -Message ""
    Write-Log -Message "Check $checkNumber`: Certificate validity (not expired)"
    Write-Log -Message "  NotBefore: $($cert.NotBefore)"
    Write-Log -Message "  NotAfter: $($cert.NotAfter)"
    Write-Log -Message "  Current time: $(Get-Date)"

    if ($cert.NotAfter -gt (Get-Date)) {
        $daysValid = [math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
        Write-Log -Message "  [PASS] Certificate is valid ($daysValid days remaining)"
    }
    else {
        Write-Log -Message "  [FAIL] Certificate is EXPIRED" -Level ERROR
        $allPassed = $false
    }

    # Check 5: NotAfter is later than previous cert (if applicable)
    if ($null -ne $PreviousNotAfter) {
        $checkNumber++
        Write-Log -Message ""
        Write-Log -Message "Check $checkNumber`: NotAfter improvement over previous certificate"
        Write-Log -Message "  Previous NotAfter: $PreviousNotAfter"
        Write-Log -Message "  New NotAfter: $($cert.NotAfter)"

        if ($cert.NotAfter -gt $PreviousNotAfter) {
            $improvement = [math]::Floor(($cert.NotAfter - $PreviousNotAfter).TotalDays)
            Write-Log -Message "  [PASS] New certificate extends validity by $improvement days"
        }
        else {
            Write-Log -Message "  [FAIL] New certificate does NOT extend validity" -Level ERROR
            $allPassed = $false
        }
    }

    # Summary
    Write-Log -Message ""
    Write-Log -Message "=" * 50
    if ($allPassed) {
        Write-Log -Message "VERIFICATION RESULT: ALL CHECKS PASSED"
        Write-Log -Message "Certificate is compliant and ready for use"
    }
    else {
        Write-Log -Message "VERIFICATION RESULT: ONE OR MORE CHECKS FAILED" -Level ERROR
        Write-Log -Message "Certificate may not function correctly for LDAPS" -Level ERROR
    }
    Write-Log -Message "=" * 50

    return $allPassed
}
#endregion

#region Cleanup Functions
function Remove-SupersededCertificates {
    <#
    .SYNOPSIS
        Removes older LDAPS certificates after successful enrollment.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$NewThumbprint,

        [Parameter(Mandatory = $true)]
        [array]$AllCandidates
    )

    Write-LogSection -Title "Certificate Cleanup"

    Write-Log -Message "Cleanup requested - evaluating superseded certificates..."
    Write-Log -Message "New certificate thumbprint (to keep): $NewThumbprint"

    $toRemove = $AllCandidates | Where-Object { $_.Thumbprint -ne $NewThumbprint }

    if ($toRemove.Count -eq 0) {
        Write-Log -Message "No superseded certificates to remove"
        return
    }

    Write-Log -Message "Found $($toRemove.Count) superseded certificate(s):"
    foreach ($candidate in $toRemove) {
        Write-Log -Message "  - $($candidate.Thumbprint)"
        Write-Log -Message "    Subject: $($candidate.Subject)"
        Write-Log -Message "    NotAfter: $($candidate.NotAfter)"
        Write-Log -Message "    Status: $(if ($candidate.IsExpired) { 'Expired' } else { 'Valid' })"
    }

    Write-Log -Message ""

    foreach ($candidate in $toRemove) {
        $certPath = "Cert:\LocalMachine\My\$($candidate.Thumbprint)"

        if ($PSCmdlet.ShouldProcess($candidate.Thumbprint, "Remove superseded LDAPS certificate")) {
            try {
                Write-Log -Message "Removing certificate: $($candidate.Thumbprint)"
                Remove-Item -Path $certPath -Force
                Write-Log -Message "  Successfully removed"
            }
            catch {
                Write-Log -Message "  Failed to remove: $_" -Level WARN
            }
        }
        else {
            Write-Log -Message "[WhatIf] Would remove: $($candidate.Thumbprint)"
        }
    }
}
#endregion

#region Main Execution
function Invoke-LdapsCertRenewal {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-LogSection -Title "LDAPS Certificate Renewal - Started"

    Write-Log -Message "Script version: $script:ScriptVersion"
    Write-Log -Message "Script start time: $($script:ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
    Write-Log -Message "Verbose logging: $script:VerboseEnabled"
    Write-Log -Message "WhatIf mode: $WhatIfPreference"

    # Log all input parameters
    Write-Log -Message ""
    Write-Log -Message "Input Parameters:"
    Write-Log -Message "  CAConfig: $(if ([string]::IsNullOrWhiteSpace($CAConfig)) { '(auto-discover)' } else { $CAConfig })"
    Write-Log -Message "  PreferredCA: $(if ([string]::IsNullOrWhiteSpace($PreferredCA)) { '(not specified)' } else { $PreferredCA })"
    Write-Log -Message "  TemplateName: $TemplateName"
    Write-Log -Message "  BaseDomain: $(if ([string]::IsNullOrWhiteSpace($BaseDomain)) { '(not specified)' } else { $BaseDomain })"
    Write-Log -Message "  IncludeShortNameSan: $IncludeShortNameSan"
    Write-Log -Message "  RenewWithinDays: $RenewWithinDays"
    Write-Log -Message "  LogPath: $LogPath"
    Write-Log -Message "  CleanupOld: $CleanupOld"
    Write-Log -Message "  MinKeySize: $MinKeySize"
    Write-Log -Message "  HashAlgorithm: $HashAlgorithm"

    try {
        # Get environment info
        Get-EnvironmentInfo

        # Resolve CA configuration (explicit or auto-discover)
        $resolvedCAConfig = Resolve-CAConfiguration -ExplicitCAConfig $CAConfig `
            -PreferredCA $PreferredCA -TemplateName $TemplateName

        # Get DC identity
        $dcIdentity = Get-DcIdentity

        # Discover candidate certificates
        $candidates = @(Get-LdapsCandidateCertificates -DcFqdn $dcIdentity.FQDN)

        # Determine state and required action
        $state = Get-CertificateState -Candidates $candidates -RenewThresholdDays $RenewWithinDays `
            -RequiredBaseDomain $BaseDomain -DcFqdn $dcIdentity.FQDN `
            -RequireShortName $IncludeShortNameSan -DcHostname $dcIdentity.Hostname

        # Log state summary
        Write-Log -Message ""
        Write-Log -Message "State Summary:"
        Write-Log -Message "  State: $($state.State)"
        Write-Log -Message "  Description: $($state.Description)"
        Write-Log -Message "  Action Required: $($state.ActionRequired)"
        Write-Log -Message "  Reason: $($state.Reason)"

        # Exit if no action required
        if (-not $state.ActionRequired) {
            Write-Log -Message ""
            Write-Log -Message "No certificate renewal required - exiting with success"

            $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
            Write-LogSection -Title "LDAPS Certificate Renewal - Completed (No Action)"
            Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
            Write-Log -Message "Exit code: 0"

            return 0
        }

        # Record previous NotAfter for comparison
        $previousNotAfter = $null
        if ($null -ne $state.BestCandidate) {
            $previousNotAfter = $state.BestCandidate.NotAfter
            Write-Log -Message ""
            Write-Log -Message "Previous certificate NotAfter: $previousNotAfter"
        }

        # Perform enrollment
        $enrollResult = Request-LdapsCertificate -CAConfig $resolvedCAConfig -TemplateName $TemplateName `
            -DcFqdn $dcIdentity.FQDN -DcHostname $dcIdentity.Hostname `
            -BaseDomain $BaseDomain -IncludeShortName $IncludeShortNameSan `
            -KeySize $MinKeySize -HashAlgorithm $HashAlgorithm

        if ($enrollResult.WhatIf) {
            Write-Log -Message ""
            Write-Log -Message "WhatIf mode - no changes were made"

            $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
            Write-LogSection -Title "LDAPS Certificate Renewal - Completed (WhatIf)"
            Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
            Write-Log -Message "Exit code: 0"

            return 0
        }

        if ($enrollResult.Pending) {
            Write-Log -Message ""
            Write-Log -Message "Certificate request is pending CA manager approval" -Level WARN
            Write-Log -Message "Request ID: $($enrollResult.RequestId)" -Level WARN

            $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
            Write-LogSection -Title "LDAPS Certificate Renewal - Pending CA Approval"
            Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
            Write-Log -Message "Exit code: 2 (pending)"

            return 2  # Special exit code for pending
        }

        if (-not $enrollResult.Success) {
            Write-Log -Message ""
            Write-Log -Message "Certificate enrollment failed" -Level ERROR

            $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
            Write-LogSection -Title "LDAPS Certificate Renewal - Failed"
            Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
            Write-Log -Message "Exit code: 1 (error)"

            return 1
        }

        # Verify new certificate
        $verifyPassed = Test-NewCertificateCompliance -Thumbprint $enrollResult.Thumbprint `
            -DcFqdn $dcIdentity.FQDN -BaseDomain $BaseDomain `
            -RequireShortName $IncludeShortNameSan -DcHostname $dcIdentity.Hostname `
            -PreviousNotAfter $previousNotAfter

        if (-not $verifyPassed) {
            Write-Log -Message ""
            Write-Log -Message "Certificate verification failed - cleanup will NOT be performed" -Level ERROR
            Write-Log -Message "Manual investigation required" -Level ERROR

            $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
            Write-LogSection -Title "LDAPS Certificate Renewal - Verification Failed"
            Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
            Write-Log -Message "Exit code: 1 (verification failed)"

            return 1
        }

        # Output thumbprint for rollback purposes
        Write-Log -Message ""
        Write-Log -Message "=" * 50
        Write-Log -Message "NEW CERTIFICATE THUMBPRINT: $($enrollResult.Thumbprint)"
        Write-Log -Message "=" * 50
        Write-Output "NewCertificateThumbprint: $($enrollResult.Thumbprint)"

        # Cleanup old certificates if requested
        if ($CleanupOld) {
            # Re-discover to include newly installed cert
            $allCandidates = @(Get-LdapsCandidateCertificates -DcFqdn $dcIdentity.FQDN)
            Remove-SupersededCertificates -NewThumbprint $enrollResult.Thumbprint -AllCandidates $allCandidates
        }
        else {
            Write-Log -Message ""
            Write-Log -Message "Cleanup not requested (-CleanupOld not specified)"
            Write-Log -Message "Previous certificates retained in store"
        }

        $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
        Write-LogSection -Title "LDAPS Certificate Renewal - Completed Successfully"
        Write-Log -Message "New certificate thumbprint: $($enrollResult.Thumbprint)"
        Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
        Write-Log -Message "Exit code: 0 (success)"

        return 0
    }
    catch {
        Write-Log -Message "" -Level ERROR
        Write-Log -Message "FATAL ERROR ENCOUNTERED" -Level ERROR
        Write-Log -Message "Error message: $_" -Level ERROR
        Write-Log -Message "Error type: $($_.Exception.GetType().FullName)" -Level ERROR

        if ($_.ScriptStackTrace) {
            Write-Log -Message "" -Level ERROR
            Write-Log -Message "Stack trace:" -Level ERROR
            foreach ($line in ($_.ScriptStackTrace -split "`n")) {
                Write-Log -Message "  $line" -Level ERROR
            }
        }

        if ($_.Exception.InnerException) {
            Write-Log -Message "" -Level ERROR
            Write-Log -Message "Inner exception: $($_.Exception.InnerException.Message)" -Level ERROR
        }

        $elapsed = ((Get-Date) - $script:ScriptStartTime).TotalSeconds
        Write-LogSection -Title "LDAPS Certificate Renewal - Failed with Exception"
        Write-Log -Message "Total execution time: $([math]::Round($elapsed, 2)) seconds"
        Write-Log -Message "Exit code: 1 (exception)"

        return 1
    }
}

# Initialize logging and run
Initialize-Logging
$script:ExitCode = Invoke-LdapsCertRenewal
exit $script:ExitCode
#endregion
