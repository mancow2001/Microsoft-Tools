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
$appScopingReport = @()

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
    
    # Build application scoping report
    if ($policy.State -ne 'disabled') {
        if ($policy.Conditions.Applications.IncludeApplications) {
            foreach ($appId in $policy.Conditions.Applications.IncludeApplications) {
                $appName = switch ($appId) {
                    "All" { "All cloud apps" }
                    "Office365" { "Office 365" }
                    "MicrosoftAdminPortals" { "Microsoft Admin Portals" }
                    default {
                        try {
                            $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
                            if ($sp) { $sp.DisplayName } else { "Unknown App" }
                        } catch { "Unknown App" }
                    }
                }
                
                $appScopingReport += [PSCustomObject]@{
                    'Policy Name' = $policy.DisplayName
                    'Policy State' = $policy.State
                    'Application Name' = $appName
                    'Application ID' = $appId
                    'Scope Type' = 'Included'
                    'Policy ID' = $policy.Id
                }
            }
        }
        
        if ($policy.Conditions.Applications.ExcludeApplications) {
            foreach ($appId in $policy.Conditions.Applications.ExcludeApplications) {
                $appName = switch ($appId) {
                    "All" { "All cloud apps" }
                    "Office365" { "Office 365" }
                    "MicrosoftAdminPortals" { "Microsoft Admin Portals" }
                    default {
                        try {
                            $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
                            if ($sp) { $sp.DisplayName } else { "Unknown App" }
                        } catch { "Unknown App" }
                    }
                }
                
                $appScopingReport += [PSCustomObject]@{
                    'Policy Name' = $policy.DisplayName
                    'Policy State' = $policy.State
                    'Application Name' = $appName
                    'Application ID' = $appId
                    'Scope Type' = 'Excluded'
                    'Policy ID' = $policy.Id
                }
            }
        }
    }
}

#------------------------------------------------------------------------------
# Build App-to-Policy Mapping
#------------------------------------------------------------------------------
$appPolicyMapping = @()
$allTargetedApps = @{}
foreach ($p in $policies) {
    if ($p.State -ne 'disabled' -and $p.Conditions.Applications.IncludeApplications) {
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
$appPolicyMapping = $appPolicyMapping | Sort-Object -Property 'Number of Policies' -Descending

#------------------------------------------------------------------------------
# EXPORT SECTION
#------------------------------------------------------------------------------
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$excelFileName     = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.xlsx"
$csvFileName       = Join-Path $ExportPath "ConditionalAccess-Policies-$timestamp.csv"
$jsonFileName      = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.json"
$htmlFileName      = Join-Path $ExportPath "ConditionalAccess-Report-$timestamp.html"
$locationsCSV      = Join-Path $ExportPath "ConditionalAccess-NamedLocations-$timestamp.csv"
$appScopingCSV     = Join-Path $ExportPath "ConditionalAccess-ApplicationScoping-$timestamp.csv"
$appMappingCSV     = Join-Path $ExportPath "ConditionalAccess-ApplicationMapping-$timestamp.csv"

Write-Host "`nExporting reports..." -ForegroundColor Yellow

switch ($ExportFormat) {
    "CSV" {
        $detailedReport | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8
        $namedLocationsReport | Export-Csv -Path $locationsCSV -NoTypeInformation -Encoding UTF8
        $appScopingReport | Export-Csv -Path $appScopingCSV -NoTypeInformation -Encoding UTF8
        $appPolicyMapping | Export-Csv -Path $appMappingCSV -NoTypeInformation -Encoding UTF8
        Write-Host "CSV files exported to: $ExportPath" -ForegroundColor Green
    }
    "JSON" {
        $completeReport = @{
            'ExportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            'Summary' = @{
                'TotalPolicies' = $policies.Count
                'EnabledPolicies' = ($policies | Where-Object {$_.State -eq 'enabled'}).Count
                'NamedLocations' = $namedLocations.Count
                'UniqueApplications' = $appPolicyMapping.Count
            }
            'ConditionalAccessPolicies' = $detailedReport
            'NamedLocations' = $namedLocationsReport
            'ApplicationScoping' = $appScopingReport
            'ApplicationPolicyMapping' = $appPolicyMapping
        }
        $completeReport | ConvertTo-Json -Depth 10 | Out-File $jsonFileName -Encoding UTF8
        Write-Host "JSON exported to: $jsonFileName" -ForegroundColor Green
    }
    "Excel" {
        # Create Summary Data
        $summaryData = @(
            [PSCustomObject]@{'Metric' = 'Total Policies'; 'Count' = $policies.Count}
            [PSCustomObject]@{'Metric' = 'Enabled Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'enabled'}).Count}
            [PSCustomObject]@{'Metric' = 'Disabled Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'disabled'}).Count}
            [PSCustomObject]@{'Metric' = 'Report Only Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'enabledForReportingButNotEnforced'}).Count}
            [PSCustomObject]@{'Metric' = 'Total Named Locations'; 'Count' = $namedLocations.Count}
            [PSCustomObject]@{'Metric' = 'Unique Applications Targeted'; 'Count' = $appPolicyMapping.Count}
            [PSCustomObject]@{'Metric' = 'Total App-Policy Assignments'; 'Count' = $appScopingReport.Count}
        )
        
        $summaryData | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "Summary"
        $detailedReport | Export-Excel -Path $excelFileName -WorksheetName "CA Policies" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "CAPolicies"
        $namedLocationsReport | Export-Excel -Path $excelFileName -WorksheetName "Named Locations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "NamedLocations"
        $appScopingReport | Export-Excel -Path $excelFileName -WorksheetName "Application Scoping" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "AppScoping"
        $appPolicyMapping | Export-Excel -Path $excelFileName -WorksheetName "App-Policy Mapping" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "AppPolicyMapping"
        Write-Host "Excel workbook exported to: $excelFileName" -ForegroundColor Green
    }
    "HTML" {
        # Build interactive HTML report
        $topApps = $appPolicyMapping | Select-Object -First 5
        $topAppsHtml = ""
        foreach ($app in $topApps) {
            $topAppsHtml += "<div class='app-item'><span>$($app.'Application Name')</span><strong>$($app.'Number of Policies') policies</strong></div>"
        }
        
        # Build policy table rows
        $policiesRows = ""
        foreach ($p in $detailedReport) {
            $badge = switch ($p.State) {
                "enabled" { "<span class='badge badge-enabled'>Enabled</span>" }
                "disabled" { "<span class='badge badge-disabled'>Disabled</span>" }
                default { "<span class='badge badge-report'>Report Only</span>" }
            }
            $policiesRows += "<tr><td>$($p.'Policy Name')</td><td>$badge</td><td>$($p.'Included Applications')</td><td>$($p.'Grant Controls')</td><td>$($p.'Included Users')</td><td>$($p.'Included Groups')</td></tr>"
        }
        
        # Build locations table rows
        $locationsRows = ""
        foreach ($loc in $namedLocationsReport) {
            $trusted = if ($loc.'Is Trusted' -eq $true) { "<span class='badge badge-trusted'>Trusted</span>" } else { "" }
            $locationsRows += "<tr><td>$($loc.'Location Name')</td><td>$($loc.'Location Type')</td><td>$trusted</td><td>$($loc.'IP/CIDR Range')</td><td>$($loc.'Country/Region')</td></tr>"
        }
        
        # Build app scoping rows
        $appScopingRows = ""
        foreach ($app in $appScopingReport) {
            $scopeBadge = if ($app.'Scope Type' -eq "Included") { "<span class='badge badge-included'>Included</span>" } else { "<span class='badge badge-excluded'>Excluded</span>" }
            $appScopingRows += "<tr><td>$($app.'Policy Name')</td><td>$($app.'Application Name')</td><td>$scopeBadge</td><td>$($app.'Application ID')</td></tr>"
        }
        
        # Build app mapping rows
        $appMappingRows = ""
        foreach ($mapping in $appPolicyMapping) {
            $appMappingRows += "<tr><td>$($mapping.'Application Name')</td><td>$($mapping.'Number of Policies')</td><td>$($mapping.'Policy Names')</td><td>$($mapping.'Application ID')</td></tr>"
        }

$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conditional Access Policy Report - $(Get-Date -Format "yyyy-MM-dd")</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .tabs { display: flex; background: #f5f5f5; border-bottom: 2px solid #ddd; overflow-x: auto; }
        .tab { padding: 15px 30px; cursor: pointer; border: none; background: transparent; font-size: 1em; font-weight: 500; color: #666; transition: all 0.3s; white-space: nowrap; }
        .tab:hover { background: #e0e0e0; color: #333; }
        .tab.active { background: white; color: #667eea; border-bottom: 3px solid #667eea; }
        .tab-content { display: none; padding: 30px; animation: fadeIn 0.5s; }
        .tab-content.active { display: block; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .summary-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3); }
        .summary-card h3 { font-size: 0.9em; opacity: 0.9; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px; }
        .summary-card .value { font-size: 2.5em; font-weight: bold; }
        .search-box { width: 100%; padding: 12px 20px; margin-bottom: 20px; border: 2px solid #ddd; border-radius: 5px; font-size: 1em; transition: border-color 0.3s; }
        .search-box:focus { outline: none; border-color: #667eea; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; background: white; box-shadow: 0 2px 10px rgba(0,0,0,0.1); border-radius: 5px; overflow: hidden; }
        thead { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        th { padding: 15px; text-align: left; font-weight: 600; cursor: pointer; user-select: none; position: relative; }
        th:hover { background: rgba(255,255,255,0.1); }
        th.sortable::after { content: ' ⇅'; opacity: 0.5; }
        td { padding: 12px 15px; border-bottom: 1px solid #f0f0f0; }
        tr:hover { background: #f9f9f9; }
        .badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 0.85em; font-weight: 500; }
        .badge-enabled { background: #d4edda; color: #155724; }
        .badge-disabled { background: #f8d7da; color: #721c24; }
        .badge-report { background: #d1ecf1; color: #0c5460; }
        .badge-trusted { background: #d4edda; color: #155724; }
        .badge-included { background: #cce5ff; color: #004085; }
        .badge-excluded { background: #f8d7da; color: #721c24; }
        .top-apps { background: #f9f9f9; padding: 20px; border-radius: 5px; margin-top: 20px; }
        .top-apps h3 { margin-bottom: 15px; color: #667eea; }
        .app-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #ddd; }
        .app-item:last-child { border-bottom: none; }
        .footer { text-align: center; padding: 20px; background: #f5f5f5; color: #666; font-size: 0.9em; }
        @media (max-width: 768px) { .summary-grid { grid-template-columns: 1fr; } table { font-size: 0.9em; } th, td { padding: 10px; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔐 Conditional Access Policy Report</h1>
            <p>Generated on $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss")</p>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="openTab(event, 'summary')">📊 Summary</button>
            <button class="tab" onclick="openTab(event, 'policies')">🛡️ CA Policies</button>
            <button class="tab" onclick="openTab(event, 'locations')">📍 Named Locations</button>
            <button class="tab" onclick="openTab(event, 'appscoping')">🎯 Application Scoping</button>
            <button class="tab" onclick="openTab(event, 'appmapping')">🔗 App-Policy Mapping</button>
        </div>
        
        <div id="summary" class="tab-content active">
            <div class="summary-grid">
                <div class="summary-card">
                    <h3>Total Policies</h3>
                    <div class="value">$($policies.Count)</div>
                </div>
                <div class="summary-card">
                    <h3>Enabled Policies</h3>
                    <div class="value">$(($policies | Where-Object {$_.State -eq 'enabled'}).Count)</div>
                </div>
                <div class="summary-card">
                    <h3>Named Locations</h3>
                    <div class="value">$($namedLocations.Count)</div>
                </div>
                <div class="summary-card">
                    <h3>Unique Applications</h3>
                    <div class="value">$($appPolicyMapping.Count)</div>
                </div>
            </div>
            
            <div class="top-apps">
                <h3>Top 5 Most Targeted Applications</h3>
                $topAppsHtml
            </div>
        </div>
        
        <div id="policies" class="tab-content">
            <input type="text" class="search-box" id="searchPolicies" placeholder="🔍 Search policies..." onkeyup="searchTable('searchPolicies', 'policiesTable')">
            <table id="policiesTable">
                <thead>
                    <tr>
                        <th class="sortable" onclick="sortTable('policiesTable', 0)">Policy Name</th>
                        <th class="sortable" onclick="sortTable('policiesTable', 1)">State</th>
                        <th class="sortable" onclick="sortTable('policiesTable', 2)">Included Applications</th>
                        <th class="sortable" onclick="sortTable('policiesTable', 3)">Grant Controls</th>
                        <th class="sortable" onclick="sortTable('policiesTable', 4)">Included Users</th>
                        <th class="sortable" onclick="sortTable('policiesTable', 5)">Included Groups</th>
                    </tr>
                </thead>
                <tbody>
                    $policiesRows
                </tbody>
            </table>
        </div>
        
        <div id="locations" class="tab-content">
            <input type="text" class="search-box" id="searchLocations" placeholder="🔍 Search locations..." onkeyup="searchTable('searchLocations', 'locationsTable')">
            <table id="locationsTable">
                <thead>
                    <tr>
                        <th class="sortable" onclick="sortTable('locationsTable', 0)">Location Name</th>
                        <th class="sortable" onclick="sortTable('locationsTable', 1)">Type</th>
                        <th>Trusted</th>
                        <th class="sortable" onclick="sortTable('locationsTable', 3)">IP/CIDR Range</th>
                        <th class="sortable" onclick="sortTable('locationsTable', 4)">Country/Region</th>
                    </tr>
                </thead>
                <tbody>
                    $locationsRows
                </tbody>
            </table>
        </div>
        
        <div id="appscoping" class="tab-content">
            <input type="text" class="search-box" id="searchAppScoping" placeholder="🔍 Search application scoping..." onkeyup="searchTable('searchAppScoping', 'appScopingTable')">
            <table id="appScopingTable">
                <thead>
                    <tr>
                        <th class="sortable" onclick="sortTable('appScopingTable', 0)">Policy Name</th>
                        <th class="sortable" onclick="sortTable('appScopingTable', 1)">Application Name</th>
                        <th class="sortable" onclick="sortTable('appScopingTable', 2)">Scope Type</th>
                        <th>Application ID</th>
                    </tr>
                </thead>
                <tbody>
                    $appScopingRows
                </tbody>
            </table>
        </div>
        
        <div id="appmapping" class="tab-content">
            <input type="text" class="search-box" id="searchAppMapping" placeholder="🔍 Search app-policy mapping..." onkeyup="searchTable('searchAppMapping', 'appMappingTable')">
            <table id="appMappingTable">
                <thead>
                    <tr>
                        <th class="sortable" onclick="sortTable('appMappingTable', 0)">Application Name</th>
                        <th class="sortable" onclick="sortTable('appMappingTable', 1)">Number of Policies</th>
                        <th>Policy Names</th>
                        <th>Application ID</th>
                    </tr>
                </thead>
                <tbody>
                    $appMappingRows
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Generated by Conditional Access Policy Report Script | Microsoft Entra ID</p>
        </div>
    </div>
    
    <script>
        function openTab(evt, tabName) {
            var i, tabcontent, tabs;
            
            tabcontent = document.getElementsByClassName("tab-content");
            for (i = 0; i < tabcontent.length; i++) {
                tabcontent[i].classList.remove("active");
            }
            
            tabs = document.getElementsByClassName("tab");
            for (i = 0; i < tabs.length; i++) {
                tabs[i].classList.remove("active");
            }
            
            document.getElementById(tabName).classList.add("active");
            evt.currentTarget.classList.add("active");
        }
        
        function searchTable(inputId, tableId) {
            var input, filter, table, tr, td, i, j, txtValue;
            input = document.getElementById(inputId);
            filter = input.value.toUpperCase();
            table = document.getElementById(tableId);
            tr = table.getElementsByTagName("tr");
            
            for (i = 1; i < tr.length; i++) {
                tr[i].style.display = "none";
                td = tr[i].getElementsByTagName("td");
                for (j = 0; j < td.length; j++) {
                    if (td[j]) {
                        txtValue = td[j].textContent || td[j].innerText;
                        if (txtValue.toUpperCase().indexOf(filter) > -1) {
                            tr[i].style.display = "";
                            break;
                        }
                    }
                }
            }
        }
        
        function sortTable(tableId, column) {
            var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
            table = document.getElementById(tableId);
            switching = true;
            dir = "asc";
            
            while (switching) {
                switching = false;
                rows = table.rows;
                
                for (i = 1; i < (rows.length - 1); i++) {
                    shouldSwitch = false;
                    x = rows[i].getElementsByTagName("TD")[column];
                    y = rows[i + 1].getElementsByTagName("TD")[column];
                    
                    var xContent = x.textContent || x.innerText;
                    var yContent = y.textContent || y.innerText;
                    
                    if (dir == "asc") {
                        if (xContent.toLowerCase() > yContent.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    } else if (dir == "desc") {
                        if (xContent.toLowerCase() < yContent.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
                
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                    switchcount++;
                } else {
                    if (switchcount == 0 && dir == "asc") {
                        dir = "desc";
                        switching = true;
                    }
                }
            }
        }
    </script>
</body>
</html>
"@
        
        $htmlContent | Out-File -FilePath $htmlFileName -Encoding UTF8
        Write-Host "HTML report exported to: $htmlFileName" -ForegroundColor Green
    }
    "All" {
        # Export CSV
        $detailedReport | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8
        $namedLocationsReport | Export-Csv -Path $locationsCSV -NoTypeInformation -Encoding UTF8
        $appScopingReport | Export-Csv -Path $appScopingCSV -NoTypeInformation -Encoding UTF8
        $appPolicyMapping | Export-Csv -Path $appMappingCSV -NoTypeInformation -Encoding UTF8
        
        # Export JSON
        $completeReport = @{
            'ExportDate' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            'Summary' = @{
                'TotalPolicies' = $policies.Count
                'EnabledPolicies' = ($policies | Where-Object {$_.State -eq 'enabled'}).Count
                'NamedLocations' = $namedLocations.Count
                'UniqueApplications' = $appPolicyMapping.Count
            }
            'ConditionalAccessPolicies' = $detailedReport
            'NamedLocations' = $namedLocationsReport
            'ApplicationScoping' = $appScopingReport
            'ApplicationPolicyMapping' = $appPolicyMapping
        }
        $completeReport | ConvertTo-Json -Depth 10 | Out-File $jsonFileName -Encoding UTF8
        
        # Export Excel
        $summaryData = @(
            [PSCustomObject]@{'Metric' = 'Total Policies'; 'Count' = $policies.Count}
            [PSCustomObject]@{'Metric' = 'Enabled Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'enabled'}).Count}
            [PSCustomObject]@{'Metric' = 'Disabled Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'disabled'}).Count}
            [PSCustomObject]@{'Metric' = 'Report Only Policies'; 'Count' = ($policies | Where-Object {$_.State -eq 'enabledForReportingButNotEnforced'}).Count}
            [PSCustomObject]@{'Metric' = 'Total Named Locations'; 'Count' = $namedLocations.Count}
            [PSCustomObject]@{'Metric' = 'Unique Applications Targeted'; 'Count' = $appPolicyMapping.Count}
            [PSCustomObject]@{'Metric' = 'Total App-Policy Assignments'; 'Count' = $appScopingReport.Count}
        )
        
        $summaryData | Export-Excel -Path $excelFileName -WorksheetName "Summary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "Summary"
        $detailedReport | Export-Excel -Path $excelFileName -WorksheetName "CA Policies" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "CAPolicies"
        $namedLocationsReport | Export-Excel -Path $excelFileName -WorksheetName "Named Locations" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "NamedLocations"
        $appScopingReport | Export-Excel -Path $excelFileName -WorksheetName "Application Scoping" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "AppScoping"
        $appPolicyMapping | Export-Excel -Path $excelFileName -WorksheetName "App-Policy Mapping" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "AppPolicyMapping"
        
        # Export HTML (reusing the same HTML content from above)
        $htmlContent | Out-File -FilePath $htmlFileName -Encoding UTF8
        
        Write-Host "All report formats exported to: $ExportPath" -ForegroundColor Green
    }
}

#------------------------------------------------------------------------------
# Summary
#------------------------------------------------------------------------------
Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Total Policies: $($policies.Count)" -ForegroundColor White
Write-Host "Enabled Policies: $(($policies | Where-Object {$_.State -eq 'enabled'}).Count)" -ForegroundColor Green
Write-Host "Disabled Policies: $(($policies | Where-Object {$_.State -eq 'disabled'}).Count)" -ForegroundColor Yellow
Write-Host "Report Only Policies: $(($policies | Where-Object {$_.State -eq 'enabledForReportingButNotEnforced'}).Count)" -ForegroundColor Cyan
Write-Host "`nNamed Locations: $($namedLocationsReport.Count)" -ForegroundColor White
Write-Host "Applications Targeted: $($appPolicyMapping.Count)" -ForegroundColor White

if ($appPolicyMapping.Count -gt 0) {
    Write-Host "`nTop 5 Most Targeted Applications:" -ForegroundColor White
    $appPolicyMapping | Select-Object -First 5 | ForEach-Object {
        Write-Host "  - $($_.'Application Name'): $($_.'Number of Policies') policies" -ForegroundColor Cyan
    }
}

Write-Host "=====================================================" -ForegroundColor Cyan

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Yellow
Write-Host "Report generation complete!" -ForegroundColor Green
