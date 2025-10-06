<#
.SYNOPSIS
    Generates a comprehensive report of Conditional Access policies and their targeted Enterprise Applications.

.DESCRIPTION
    This script connects to Microsoft Graph and exports detailed information about:
    - All Conditional Access policies
    - Enterprise Applications targeted by each policy
    - Complete policy configurations including users, groups, locations, device states, and grant controls
    - Named locations referenced in policies
    - Automatically checks for and installs required Microsoft Graph modules
    - Exports all data to a single Excel workbook with multiple worksheets (or CSV/JSON/HTML)

.PARAMETER ExportPath
    Path where the report files will be saved. Defaults to current directory.

.PARAMETER ExportFormat
    Export format: Excel, CSV, JSON, HTML, or All. Defaults to Excel.

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -ExportPath "C:\Reports" -ExportFormat "Excel"
    
.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -ExportFormat "All"

.NOTES
    Requires Microsoft Graph PowerShell SDK and ImportExcel module.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Excel","CSV","JSON","HTML","All")]
    [string]$ExportFormat = "Excel"
)

#------------------------------------------------------------------------------
# Function: Install-RequiredModules
#------------------------------------------------------------------------------
function Install-RequiredModules {
    param(
        [string[]]$ModuleNames
    )

    Write-Host "`nChecking required PowerShell modules..." -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor Cyan
    
    $modulesToInstall = @()
    
    foreach ($moduleName in $ModuleNames) {
        Write-Host "Checking for module: $moduleName..." -ForegroundColor Yellow -NoNewline
        $module = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
        if ($module) {
            Write-Host " FOUND (Version: $($module.Version))" -ForegroundColor Green
        } else {
            Write-Host " NOT FOUND" -ForegroundColor Red
            $modulesToInstall += $moduleName
        }
    }

    if ($modulesToInstall.Count -gt 0) {
        Write-Host "`nThe following modules need to be installed:" -ForegroundColor Yellow
        $modulesToInstall | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
        $response = Read-Host "`nDo you want to install these modules now? (Y/N)"
        
        if ($response -match '^[Yy]$') {
            foreach ($moduleName in $modulesToInstall) {
                Write-Host "`nInstalling $moduleName..." -ForegroundColor Yellow
                try {
                    Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Host "Successfully installed $moduleName" -ForegroundColor Green
                } catch {
                    Write-Host "ERROR: Failed to install $moduleName" -ForegroundColor Red
                    Write-Host $_.Exception.Message -ForegroundColor Red
                    return $false
                }
            }
        } else {
            Write-Host "Installation cancelled. Required modules missing." -ForegroundColor Red
            return $false
        }
    } else {
        Write-Host "All required modules are already installed." -ForegroundColor Green
    }

    Write-Host "=====================================================" -ForegroundColor Cyan
    return $true
}

#------------------------------------------------------------------------------
# Required modules list
#------------------------------------------------------------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Users',
    'ImportExcel'
)

#------------------------------------------------------------------------------
# Check and install modules
#------------------------------------------------------------------------------
if (-not (Install-RequiredModules -ModuleNames $requiredModules)) {
    exit 1
}

# Import ImportExcel explicitly
Import-Module ImportExcel -ErrorAction SilentlyContinue

#------------------------------------------------------------------------------
# Helper functions
#------------------------------------------------------------------------------
function Get-GroupMemberNames {
    param($GroupIds)
    $names = @()
    foreach ($id in $GroupIds) {
        try {
            $g = Get-MgGroup -GroupId $id -ErrorAction SilentlyContinue
            if ($g) { $names += "$($g.DisplayName) ($id)" } else { $names += "Group not found ($id)" }
        } catch { $names += "Error retrieving group ($id)" }
    }
    return $names -join "; "
}

function Get-UserNames {
    param($UserIds)
    $names = @()
    foreach ($id in $UserIds) {
        try {
            $u = Get-MgUser -UserId $id -ErrorAction SilentlyContinue
            if ($u) { $names += "$($u.DisplayName) ($($u.UserPrincipalName))" } else { $names += "User not found ($id)" }
        } catch { $names += "Error retrieving user ($id)" }
    }
    return $names -join "; "
}

function Get-RoleNames {
    param($RoleIds)
    $names = @()
    foreach ($id in $RoleIds) {
        try {
            $r = Get-MgDirectoryRoleTemplate -DirectoryRoleTemplateId $id -ErrorAction SilentlyContinue
            if ($r) { $names += "$($r.DisplayName) ($id)" } else { $names += "Role: $id" }
        } catch { $names += "Role: $id" }
    }
    return $names -join "; "
}

function Get-ApplicationNames {
    param($AppIds)
    $names = @()
    foreach ($id in $AppIds) {
        switch ($id) {
            "All" { $names += "All cloud apps"; continue }
            "Office365" { $names += "Office 365"; continue }
            "MicrosoftAdminPortals" { $names += "Microsoft Admin Portals"; continue }
            default {
                try {
                    $sp = Get-MgServicePrincipal -Filter "appId eq '$id'" -ErrorAction SilentlyContinue
                    if ($sp) { $names += "$($sp.DisplayName) ($id)" } else { $names += "Unknown App ($id)" }
                } catch { $names += "Unknown App ($id)" }
            }
        }
    }
    return $names -join "; "
}

function Get-NamedLocationNames {
    param($Ids)
    $names = @()
    foreach ($id in $Ids) {
        try {
            $loc = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $id -ErrorAction SilentlyContinue
            if ($loc) { $names += "$($loc.DisplayName) ($id)" } else { $names += "Location not found ($id)" }
        } catch { $names += "Error retrieving location ($id)" }
    }
    return $names -join "; "
}

#------------------------------------------------------------------------------
# MAIN EXECUTION
#------------------------------------------------------------------------------
Write-Host "Starting Conditional Access Policy Report Generation..." -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    Write-Host "Creating export directory: $ExportPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $ExportPath | Out-Null
}

# Connect to Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "Policy.Read.All","Application.Read.All","Directory.Read.All","Group.Read.All","User.Read.All" -NoWelcome
    Write-Host "Connected successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
    exit 1
}

# Retrieve CA policies
Write-Host "`nRetrieving Conditional Access Policies..." -ForegroundColor Yellow
try {
    $policies = Get-MgIdentityConditionalAccessPolicy -All
} catch {
    Write-Host "Failed to retrieve policies: $_" -ForegroundColor Red
    Disconnect-MgGraph; exit 1
}
Write-Host "Found $($policies.Count) Conditional Access policies." -ForegroundColor Green

# Retrieve Named Locations
Write-Host "Retrieving Named Locations..." -ForegroundColor Yellow
$namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All

# Build named location data
$namedLocationsReport = @()
foreach ($loc in $namedLocations) {
    $type = $loc.AdditionalProperties['@odata.type']
    if ($type -eq '#microsoft.graph.ipNamedLocation') {
        $ranges = $loc.AdditionalProperties['ipRanges']
        if ($ranges) {
            foreach ($r in $ranges) {
                $namedLocationsReport += [PSCustomObject]@{
                    'Location Name' = $loc.DisplayName
                    'Location ID'   = $loc.Id
                    'Location Type' = 'IP Range'
                    'Is Trusted'    = $loc.AdditionalProperties['isTrusted']
                    'IP/CIDR Range' = $r['cidrAddress']
                    'Country/Region' = 'N/A'
                    'Include Unknown Countries' = 'N/A'
                    'Created' = $loc.CreatedDateTime
                    'Modified' = $loc.ModifiedDateTime
                }
            }
        }
    } elseif ($type -eq '#microsoft.graph.countryNamedLocation') {
        $countries = $loc.AdditionalProperties['countriesAndRegions'] -join ', '
        $namedLocationsReport += [PSCustomObject]@{
            'Location Name' = $loc.DisplayName
            'Location ID'   = $loc.Id
            'Location Type' = 'Country/Region'
            'Is Trusted'    = 'N/A'
            'IP/CIDR Range' = 'N/A'
            'Country/Region' = $countries
            'Include Unknown Countries' = $loc.AdditionalProperties['includeUnknownCountriesAndRegions']
            'Created' = $loc.CreatedDateTime
            'Modified' = $loc.ModifiedDateTime
        }
    }
}

#------------------------------------------------------------------------------
# Process policies into detailed report
#------------------------------------------------------------------------------
$detailedReport = @()
foreach ($policy in $policies) {
    $appsIn  = if ($policy.Conditions.Applications.IncludeApplications) { Get-ApplicationNames $policy.Conditions.Applications.IncludeApplications } else { "None" }
    $appsOut = if ($policy.Conditions.Applications.ExcludeApplications) { Get-ApplicationNames $policy.Conditions.Applications.ExcludeApplications } else { "None" }

    $usersIn  = if ($policy.Conditions.Users.IncludeUsers -contains "All") { "All users" } elseif ($policy.Conditions.Users.IncludeUsers) { Get-UserNames $policy.Conditions.Users.IncludeUsers } else { "None" }
    $usersOut = if ($policy.Conditions.Users.ExcludeUsers) { Get-UserNames $policy.Conditions.Users.ExcludeUsers } else { "None" }

    $groupsIn  = if ($policy.Conditions.Users.IncludeGroups) { Get-GroupMemberNames $policy.Conditions.Users.IncludeGroups } else { "None" }
    $groupsOut = if ($policy.Conditions.Users.ExcludeGroups) { Get-GroupMemberNames $policy.Conditions.Users.ExcludeGroups } else { "None" }

    $rolesIn  = if ($policy.Conditions.Users.IncludeRoles) { Get-RoleNames $policy.Conditions.Users.IncludeRoles } else { "None" }
    $rolesOut = if ($policy.Conditions.Users.ExcludeRoles) { Get-RoleNames $policy.Conditions.Users.ExcludeRoles } else { "None" }

    $locsIn = if ($policy.Conditions.Locations.IncludeLocations -contains "All") { "All locations" } elseif ($policy.Conditions.Locations.IncludeLocations) { Get-NamedLocationNames $policy.Conditions.Locations.IncludeLocations } else { "None" }
    $locsOut = if ($policy.Conditions.Locations.ExcludeLocations -contains "AllTrusted") { "All trusted locations" } elseif ($policy.Conditions.Locations.ExcludeLocations) { Get-NamedLocationNames $policy.Conditions.Locations.ExcludeLocations } else { "None" }

    $platIn  = if ($policy.Conditions.Platforms.IncludePlatforms) { $policy.Conditions.Platforms.IncludePlatforms -join ", " } else { "None" }
    $platOut = if ($policy.Conditions.Platforms.ExcludePlatforms) { $policy.Conditions.Platforms.ExcludePlatforms -join ", " } else { "None" }

    $grant = if ($policy.GrantControls) { ($policy.GrantControls.BuiltInControls -join ", ") } else { "None" }
    $operator = if ($policy.GrantControls.Operator) { $policy.GrantControls.Operator } else { "N/A" }

    $detailedReport += [PSCustomObject]@{
        'Policy Name'          = $policy.DisplayName
        'Policy ID'            = $policy.Id
        'State'                = $policy.State
        'Included Applications'= $appsIn
        'Excluded Applications'= $appsOut
        'Included Users'       = $usersIn
        'Excluded Users'       = $usersOut
        'Included Groups'      = $groupsIn
        'Excluded Groups'      = $groupsOut
        'Included Roles'       = $rolesIn
        'Excluded Roles'       = $rolesOut
        'Included Locations'   = $locsIn
        'Excluded Locations'   = $locsOut
        'Included Platforms'   = $platIn
        'Excluded Platforms'   = $platOut
        'Grant Controls'       = $grant
        'Grant Operator'       = $operator
        'Created'              = $policy.CreatedDateTime
        'Modified'             = $policy.ModifiedDateTime
    }
}

#------------------------------------------------------------------------------
# Build App-to-Policy Mapping
#------------------------------------------------------------------------------
$appPolicyMapping = @()
$allTargetedApps = @{}
foreach ($p in $policies) {
    if ($p.Conditions.Applications.IncludeApplications) {
        foreach ($app in $p.Conditions.Applications.IncludeApplications) {
            if (-not $allTargetedApps.ContainsKey($app)) { $allTargetedApps[$app] = @() }
            $allTargetedApps[$app] += $p.DisplayName
        }
    }
}

foreach ($appId in $allTargetedApps.Keys) {
    $appName = (Get-ApplicationNames @($appId))
    $appPolicyMapping += [PSCustomObject]@{
        'Application Name' = $appName
        'Application ID'   = $appId
        'Number of Policies' = $allTargetedApps[$appId].Count
        'Policy Names' = ($allTargetedApps[$appId] -join "; ")
    }
}

#------------------------------------------------------------------------------
# EXPORT SECTION
#------------------------------------------------------------------------------
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$excelFileName     = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.xlsx"
$csvFileName       = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.csv"
$jsonFileName      = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.json"
$htmlFileName      = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.html"
$locationsCSV      = Join-Path $ExportPath "NamedLocations-Report-$timestamp.csv"
$locationsJSON     = Join-Path $ExportPath "NamedLocations-Report-$timestamp.json"

Write-Host "`nExporting reports..." -ForegroundColor Yellow

switch ($ExportFormat) {
    "CSV" {
        $detailedReport | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8
        $namedLocationsReport | Export-Csv -Path $locationsCSV -NoTypeInformation -Encoding UTF8
        Write-Host "CSV exported to: $csvFileName" -ForegroundColor Green
    }
    "JSON" {
        $detailedReport | ConvertTo-Json -Depth 10 | Out-File $jsonFileName -Encoding UTF8
        $namedLocationsReport | ConvertTo-Json -Depth 10 | Out-File $locationsJSON -Encoding UTF8
        Write-Host "JSON exported to: $jsonFileName" -ForegroundColor Green
    }
    "Excel" {
        $detailedReport | Export-Excel -Path $excelFileName -WorksheetName "Policies" -AutoSize
        $namedLocationsReport | Export-Excel -Path $excelFileName -WorksheetName "NamedLocations" -AutoSize -Append
        $appPolicyMapping | Export-Excel -Path $excelFileName -WorksheetName "AppPolicyMapping" -AutoSize -Append
        Write-Host "Excel workbook exported to: $excelFileName" -ForegroundColor Green
    }
    "HTML" {
        $detailedReport | ConvertTo-Html -Title "Conditional Access Report" | Out-File $htmlFileName -Encoding UTF8
        Write-Host "HTML exported to: $htmlFileName" -ForegroundColor Green
    }
    "All" {
        # Export everything
        $detailedReport | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8
        $namedLocationsReport | Export-Csv -Path $locationsCSV -NoTypeInformation -Encoding UTF8
        $detailedReport | ConvertTo-Json -Depth 10 | Out-File $jsonFileName -Encoding UTF8
        $namedLocationsReport | ConvertTo-Json -Depth 10 | Out-File $locationsJSON -Encoding UTF8
        $detailedReport | Export-Excel -Path $excelFileName -WorksheetName "Policies" -AutoSize
        $namedLocationsReport | Export-Excel -Path $excelFileName -WorksheetName "NamedLocations" -AutoSize -Append
        $appPolicyMapping | Export-Excel -Path $excelFileName -WorksheetName "AppPolicyMapping" -AutoSize -Append
        $detailedReport | ConvertTo-Html -Title "Conditional Access Report" | Out-File $htmlFileName -Encoding UTF8
        Write-Host "All report formats exported to $ExportPath" -ForegroundColor Green
    }
}

#------------------------------------------------------------------------------
# Summary
#------------------------------------------------------------------------------
Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Total Policies: $($policies.Count)" -ForegroundColor White
Write-Host "Named Locations: $($namedLocationsReport.Count)" -ForegroundColor White
Write-Host "Applications Targeted: $($appPolicyMapping.Count)" -ForegroundColor White
Write-Host "=====================================================" -ForegroundColor Cyan

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Yellow
Write-Host "Report generation complete!" -ForegroundColor Green

