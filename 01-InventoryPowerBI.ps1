<#
.SYNOPSIS
    Enumerates all Power BI items (reports, semantic models, dashboards, and
    dataflows) across workspaces in the current Power BI tenant.

.DESCRIPTION
    Connects to the Power BI service and lists every Power BI item the
    authenticated user can access, grouped by workspace. Covered item types:
    reports (standard and paginated), semantic models (datasets), dashboards,
    and dataflows.

    NOTE: Power BI content is scoped to a TENANT (not an Azure subscription).
    Items live in workspaces, so this script iterates workspaces.

    NOTE: This covers the classic Power BI item types exposed by the
    MicrosoftPowerBIMgmt module. Fabric-native items (lakehouses, warehouses,
    notebooks, pipelines, etc.) are NOT included - those require the separate
    Fabric REST API.

.PARAMETER CsvPath
    Optional path to export the results as a CSV file.

.PARAMETER IncludePersonalWorkspace
    Include the user's personal "My Workspace" in the scan.

.PARAMETER AllPersonalWorkspaces
    Include every user's personal workspace in the scan. Requires Power BI /
    Fabric admin rights (uses Organization scope, type PersonalGroup).

.PARAMETER Backup
    Download a backup of every exportable item:
      - Reports            -> .pbix
      - Paginated reports  -> .rdl
      - Dataflows          -> .json (model definition)
    Files are saved under -BackupPath in a subfolder per workspace.
    NOTE: Dashboards cannot be exported (no API); their metadata is saved as
    .json instead. Reports on a live connection / DirectQuery to a shared
    dataset, and system reports (e.g. usage metrics), cannot be exported as
    .pbix and are skipped - see the backup log for details.

.PARAMETER BackupPath
    Root folder for -Backup output. Defaults to .\PowerBIBackup_<timestamp>.

# -----------------------------------------------------------------------------
#  SETUP & USAGE
# -----------------------------------------------------------------------------
#
#  Prerequisites
#  -------------
#    1. Install PowerShell 7 (recommended; installs alongside Windows PowerShell):
#         winget install --id Microsoft.PowerShell --source winget
#       Then launch it:
#         pwsh
#
#    2. Install the Power BI module:
#         Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
#
#    3. Sign in to Power BI:
#         Connect-PowerBIServiceAccount
#
#  Usage
#  -----
#    Scan all accessible workspaces:
#         .\01-InventoryPowerBI.ps1
#       (As a Fabric administrator this returns ALL workspaces in the tenant
#        via Organization scope; otherwise only workspaces your account belongs to.)
#
#    Include your personal "My Workspace":
#         .\01-InventoryPowerBI.ps1 -IncludePersonalWorkspace
#
#    Include EVERY user's personal workspace (requires Fabric admin rights):
#         .\01-InventoryPowerBI.ps1 -AllPersonalWorkspaces
#
#    Export results to CSV:
#         .\01-InventoryPowerBI.ps1 -CsvPath .\pbi-inventory.csv
#
#    Back up all exportable items (.pbix / .rdl / dataflow .json):
#         .\01-InventoryPowerBI.ps1 -Backup
#         .\01-InventoryPowerBI.ps1 -Backup -BackupPath D:\PBIBackup
#
#    Everything together:
#         .\01-InventoryPowerBI.ps1 -IncludePersonalWorkspace -CsvPath .\pbi-inventory.csv
#
#  Notes
#  -----
#    - Organization scope (full tenant visibility) requires Power BI / Fabric
#      admin rights. A Fabric administrator can enumerate ALL workspaces in the
#      tenant through the admin APIs (Organization scope) - even ones they are
#      not a member of. This is read/enumerate visibility for governance, not
#      edit access inside each workspace.
#    - Without admin rights the script falls back to user scope and lists only
#      the workspaces your account is actually a member of.
#    - "Datasets" are the API's name for what the portal now calls
#      "semantic models" - the script labels them as such.
#    - You may see repeated MSAL warnings during sign-in, e.g.:
#         "WARNING SetAuthorityUri ... without authority type, defaulting to MsSts"
#      These are harmless diagnostic messages from the authentication library
#      and can be safely IGNORED. They do not affect the results.
# -----------------------------------------------------------------------------
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvPath,                      # Optional: export results to CSV

    [Parameter()]
    [switch]$IncludePersonalWorkspace,     # Include the signed-in user's own "My Workspace"

    [Parameter()]
    [switch]$AllPersonalWorkspaces,        # Include EVERY user's personal workspace (requires Fabric admin)

    [Parameter()]
    [switch]$Backup,                       # Download .pbix / .rdl / dataflow .json for every exportable item

    [Parameter()]
    [string]$BackupPath = ''                # Root folder for -Backup output (default: .\PowerBIBackup_<timestamp>)
)

# --- Variables -------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$results    = [System.Collections.Generic.List[object]]::new()
$backupLog  = [System.Collections.Generic.List[object]]::new()

# --- Backup root setup -----------------------------------------------------
if ($Backup) {
    if ([string]::IsNullOrWhiteSpace($BackupPath)) {
        $stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $BackupPath = Join-Path -Path (Get-Location) -ChildPath "PowerBIBackup_$stamp"
    }
    if (-not (Test-Path -LiteralPath $BackupPath)) {
        New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
    }
    Write-Host "Backup enabled. Files will be saved under: $BackupPath" -ForegroundColor Cyan
}

# Sanitizes a workspace/item name so it is safe to use as a file/folder name.
function Get-SafeName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '_' }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[{0}]" -f [Regex]::Escape($invalid)
    ($Name -replace $pattern, '_').Trim()
}

# --- Ensure module is available --------------------------------------------
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    throw "MicrosoftPowerBIMgmt module not found. Install with: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
}

# --- Verify an active session ----------------------------------------------
# Sign in beforehand with: Connect-PowerBIServiceAccount
try {
    Get-PowerBIAccessToken | Out-Null
}
catch {
    throw "Not signed in to Power BI. Run 'Connect-PowerBIServiceAccount' first, then re-run this script."
}

# --- Get workspaces --------------------------------------------------------
# -Scope Organization requires tenant admin rights; falls back to user scope.
try {
    $workspaces = Get-PowerBIWorkspace -Scope Organization -All
    Write-Host "Enumerating as tenant administrator (Organization scope)." -ForegroundColor Green
}
catch {
    Write-Host "Admin scope unavailable; using individual (user) scope." -ForegroundColor Yellow
    $workspaces = Get-PowerBIWorkspace -All
}

if ($IncludePersonalWorkspace) {
    # Add a synthetic entry representing the signed-in user's own "My Workspace"
    $workspaces = @($workspaces) + ([pscustomobject]@{ Id = $null; Name = 'My Workspace' })
}

if ($AllPersonalWorkspaces) {
    # Every user's personal workspace (type PersonalGroup). Admin-only.
    try {
        $personal = Get-PowerBIWorkspace -Scope Organization -Type PersonalGroup -All
        Write-Host ("Adding {0} personal workspace(s) from all users." -f @($personal).Count) -ForegroundColor Green
        $workspaces = @($workspaces) + $personal
    }
    catch {
        Write-Warning "Could not enumerate all personal workspaces (Fabric admin rights required): $($_.Exception.Message)"
    }
}

# De-duplicate workspaces by Id so the same workspace is not scanned twice.
# (Admin scope can return a workspace more than once.) Keep null-Id entries
# such as the synthetic "My Workspace".
$workspaces = @(
    @($workspaces | Where-Object { -not $_.Id }) +
    @($workspaces | Where-Object { $_.Id } | Sort-Object Id -Unique)
)

# --- Iterate ----------------------------------------------------------------
foreach ($ws in $workspaces) {
    $wsName = $ws.Name
    Write-Host "Scanning workspace: $wsName" -ForegroundColor Cyan

    $reportParams    = @{}
    $datasetParams   = @{}
    $dashboardParams = @{}
    $dataflowParams  = @{}
    if ($ws.Id) {
        $reportParams['WorkspaceId']    = $ws.Id
        $datasetParams['WorkspaceId']   = $ws.Id
        $dashboardParams['WorkspaceId'] = $ws.Id
        $dataflowParams['WorkspaceId']  = $ws.Id
    }

    # Prepare this workspace's backup folder (only when -Backup and we have an Id)
    $wsBackupDir = $null
    if ($Backup) {
        if ($ws.Id) {
            $wsBackupDir = Join-Path -Path $BackupPath -ChildPath (Get-SafeName $wsName)
            if (-not (Test-Path -LiteralPath $wsBackupDir)) {
                New-Item -ItemType Directory -Path $wsBackupDir -Force | Out-Null
            }
        }
        else {
            # The synthetic "My Workspace" (null Id) cannot be targeted for export.
            $backupLog.Add([pscustomobject]@{
                Workspace = $wsName; Type = '-'; Name = '-'
                Status = 'Skipped'; Detail = 'Personal "My Workspace" cannot be exported by these cmdlets.'
            })
        }
    }

    # Reports (standard + paginated; paginated have ReportType 'PaginatedReport')
    try {
        $reports = Get-PowerBIReport @reportParams
        foreach ($r in $reports) {
            $isPaginated = ($r.ReportType -eq 'PaginatedReport')
            $type = if ($isPaginated) { 'PaginatedReport' } else { 'Report' }
            $results.Add([pscustomobject]@{
                Workspace = $wsName
                Type      = $type
                Name      = $r.Name
                Id        = $r.Id
                WebUrl    = $r.WebUrl
            })

            # Backup: .pbix for standard reports, .rdl for paginated
            if ($Backup -and $wsBackupDir) {
                $ext  = if ($isPaginated) { 'rdl' } else { 'pbix' }
                $file = Join-Path $wsBackupDir ("{0}_{1}.{2}" -f (Get-SafeName $r.Name), $r.Id, $ext)
                try {
                    # Pass WorkspaceId so the report resolves correctly outside My Workspace.
                    Export-PowerBIReport -WorkspaceId $ws.Id -Id $r.Id -OutFile $file -ErrorAction Stop
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = $type; Name = $r.Name
                        Status = 'Exported'; Detail = $file
                    })
                }
                catch {
                    # Common for live-connection/DirectQuery reports and system reports.
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = $type; Name = $r.Name
                        Status = 'Failed'; Detail = $_.Exception.Message
                    })
                    Write-Warning "  Export failed: $type '$($r.Name)' - $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Could not read reports in '$wsName': $($_.Exception.Message)"
    }

    # Semantic models (datasets)
    try {
        $datasets = Get-PowerBIDataset @datasetParams
        foreach ($d in $datasets) {
            $results.Add([pscustomobject]@{
                Workspace = $wsName
                Type      = 'SemanticModel'
                Name      = $d.Name
                Id        = $d.Id
                WebUrl    = $null
            })
        }
    }
    catch {
        Write-Warning "Could not read semantic models in '$wsName': $($_.Exception.Message)"
    }

    # Dashboards
    try {
        $dashboards = Get-PowerBIDashboard @dashboardParams
        foreach ($db in $dashboards) {
            $results.Add([pscustomobject]@{
                Workspace = $wsName
                Type      = 'Dashboard'
                Name      = $db.Name
                Id        = $db.Id
                WebUrl    = $db.WebUrl
            })

            # Backup: dashboards have NO export API; save metadata + tiles as JSON.
            if ($Backup -and $wsBackupDir) {
                $file = Join-Path $wsBackupDir ("Dashboard_{0}_{1}.json" -f (Get-SafeName $db.Name), $db.Id)
                try {
                    $tilesUri = "dashboards/$($db.Id)/tiles"
                    $tiles = (Invoke-PowerBIRestMethod -Url $tilesUri -Method Get | ConvertFrom-Json).value
                    [pscustomobject]@{ Dashboard = $db; Tiles = $tiles } |
                        ConvertTo-Json -Depth 10 | Out-File -FilePath $file -Encoding UTF8
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = 'Dashboard'; Name = $db.Name
                        Status = 'MetadataOnly'; Detail = $file
                    })
                }
                catch {
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = 'Dashboard'; Name = $db.Name
                        Status = 'Failed'; Detail = $_.Exception.Message
                    })
                    Write-Warning "  Dashboard metadata failed: '$($db.Name)' - $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Could not read dashboards in '$wsName': $($_.Exception.Message)"
    }

    # Dataflows
    try {
        $dataflows = Get-PowerBIDataflow @dataflowParams
        foreach ($df in $dataflows) {
            $results.Add([pscustomobject]@{
                Workspace = $wsName
                Type      = 'Dataflow'
                Name      = $df.Name
                Id        = $df.Id
                WebUrl    = $null
            })

            # Backup: export the dataflow model definition as JSON.
            # Requires a real workspace (group) Id - personal workspace can't be targeted.
            if ($Backup -and $wsBackupDir -and $ws.Id) {
                $file = Join-Path $wsBackupDir ("Dataflow_{0}_{1}.json" -f (Get-SafeName $df.Name), $df.Id)
                try {
                    $dfUri = "groups/$($ws.Id)/dataflows/$($df.Id)"
                    $definition = Invoke-PowerBIRestMethod -Url $dfUri -Method Get
                    # Response is already JSON text; write it as-is.
                    $definition | Out-File -FilePath $file -Encoding UTF8
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = 'Dataflow'; Name = $df.Name
                        Status = 'Exported'; Detail = $file
                    })
                }
                catch {
                    $backupLog.Add([pscustomobject]@{
                        Workspace = $wsName; Type = 'Dataflow'; Name = $df.Name
                        Status = 'Failed'; Detail = $_.Exception.Message
                    })
                    Write-Warning "  Dataflow export failed: '$($df.Name)' - $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Could not read dataflows in '$wsName': $($_.Exception.Message)"
    }
}

# --- Output -----------------------------------------------------------------
# De-duplicate: an item can be returned more than once when a workspace is
# reachable through both organization and user scope. Item Id is globally
# unique, so dedupe on it.
$results = $results | Sort-Object Id -Unique

if ($results.Count -eq 0) {
    Write-Host ""
    Write-Host "No Power BI items found in the scanned workspaces." -ForegroundColor Yellow
    Write-Host "(Nothing to display. If you expected results, check your permissions or try -IncludePersonalWorkspace.)" -ForegroundColor Yellow
    return
}

$results |
    Sort-Object Workspace, Type, Name |
    Format-Table Workspace, Type, Name, Id -AutoSize

Write-Host ""
Write-Host ("Total items: {0}" -f $results.Count) -ForegroundColor Green

# Show a count for every known item type, including those with zero items.
$knownTypes = 'Report', 'PaginatedReport', 'SemanticModel', 'Dashboard', 'Dataflow'
$counts = $results | Group-Object Type -AsHashTable -AsString
foreach ($t in $knownTypes) {
    $c = if ($counts -and $counts.ContainsKey($t)) { $counts[$t].Count } else { 0 }
    $color = if ($c -eq 0) { 'DarkGray' } else { 'Green' }
    $label = if ($c -eq 0) { 'none' } else { $c }
    Write-Host ("  {0,-16} {1}" -f $t, $label) -ForegroundColor $color
}

if ($CsvPath) {
    $results | Sort-Object Workspace, Type, Name | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $CsvPath" -ForegroundColor Green
}

# --- Backup summary ---------------------------------------------------------
if ($Backup) {
    Write-Host ""
    Write-Host "Backup summary" -ForegroundColor Cyan
    Write-Host "--------------" -ForegroundColor Cyan

    if ($backupLog.Count -eq 0) {
        Write-Host "  Nothing was eligible for backup." -ForegroundColor Yellow
    }
    else {
        $backupLog |
            Group-Object Status |
            Sort-Object Name |
            ForEach-Object {
                $color = switch ($_.Name) {
                    'Exported'     { 'Green' }
                    'MetadataOnly' { 'Yellow' }
                    'Skipped'      { 'DarkGray' }
                    default        { 'Red' }      # Failed
                }
                Write-Host ("  {0,-13} {1}" -f $_.Name, $_.Count) -ForegroundColor $color
            }

        # Write a full log (CSV) into the backup root for the record.
        $logFile = Join-Path $BackupPath 'backup-log.csv'
        $backupLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "  Files saved under: $BackupPath" -ForegroundColor Green
        Write-Host "  Full log:          $logFile" -ForegroundColor Green
    }

    # --- Restorability information ------------------------------------------
    $restoreInfo = @"
POWER BI BACKUP - WHAT IS RESTORABLE
====================================
Created : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Folder  : $BackupPath

RESTORABLE (use 02-RestorePowerBI.ps1, or import manually)
  .pbix  Reports + their semantic model.
         Restore: New-PowerBIReport -Path <file>.pbix -Name <name> -Workspace <id>
         (or open in Power BI Desktop / upload to the service)
  .rdl   Paginated reports. Re-upload the .rdl to a workspace.

PARTIALLY RESTORABLE (manual steps required)
  Dataflow .json  The model definition only. Can be recreated via the import
                  API, but you must rebind data source credentials and the
                  gateway afterwards. Not handled by the restore script.

NOT RESTORABLE (reference/documentation only)
  Dashboard .json  Metadata + tiles only. There is NO dashboard import API.
                   To rebuild, recreate the dashboard in the portal and re-pin
                   each tile manually, using this file as a reference.

NOT CAPTURED BY ANY EXPORT (must be reconfigured after restore)
  - Data source credentials and gateway bindings (re-enter before refresh).
  - Sharing, permissions, and Row-Level Security (RLS) role memberships.
  - App packaging, scheduled refresh settings, and workspace settings.

SUMMARY: This is a CONTENT backup (report definitions + models), not a full
disaster-recovery snapshot of the tenant. The .pbix / .rdl files restore
cleanly; everything else needs manual work or cannot be restored.
"@

    $restoreFile = Join-Path $BackupPath 'RESTORE-NOTES.txt'
    $restoreInfo | Out-File -FilePath $restoreFile -Encoding UTF8

    Write-Host ""
    Write-Host "What is restorable from this backup:" -ForegroundColor Cyan
    Write-Host "  RESTORABLE      .pbix (reports + models), .rdl (paginated reports)" -ForegroundColor Green
    Write-Host "  PARTIAL         Dataflow .json (manual: rebind credentials/gateway)" -ForegroundColor Yellow
    Write-Host "  NOT RESTORABLE  Dashboard .json (metadata only - no import API)" -ForegroundColor DarkGray
    Write-Host "  NOT CAPTURED    credentials, permissions, RLS, refresh/app settings" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Details written to: $restoreFile" -ForegroundColor Green
    Write-Host "  Restore .pbix/.rdl with: .\02-RestorePowerBI.ps1 -BackupPath `"$BackupPath`" -TargetWorkspaceId <id>" -ForegroundColor Green
}
