<#
.SYNOPSIS
    Restores Power BI reports (.pbix) and paginated reports (.rdl) from a backup
    folder created by 01-InventoryPowerBI.ps1 into a target workspace.

.DESCRIPTION
    Walks a backup folder, finds every .pbix and .rdl file, and imports each one
    into the target workspace using New-PowerBIReport. Produces a restore log.

    WHAT THIS RESTORES
      .pbix  Reports + their embedded semantic model.
      .rdl   Paginated reports.

    WHAT THIS DOES NOT RESTORE (by design - these cannot be imported)
      Dashboard .json  Metadata only; no dashboard import API exists. Skipped.
      Dataflow  .json  Must be recreated manually (rebind credentials/gateway). Skipped.

    AFTER RESTORE you will likely still need to:
      - Re-enter data source credentials and rebind the gateway before refresh.
      - Reconfigure sharing, permissions, and Row-Level Security (RLS).
      - Re-establish scheduled refresh, app packaging, and workspace settings.
      - For reports bound to an external/shared dataset, rebind them to the
        dataset in the target workspace (the .pbix keeps its original binding).

.PARAMETER BackupPath
    Path to the backup folder produced by 01-InventoryPowerBI.ps1. The script
    searches it recursively for .pbix and .rdl files.

.PARAMETER TargetWorkspaceId
    Id of the workspace to restore into. If omitted, items are restored to the
    signed-in user's personal "My Workspace".

.PARAMETER ConflictAction
    What to do if an item with the same name already exists. One of:
    Ignore, Abort, Overwrite, CreateOrOverwrite. Default: CreateOrOverwrite.

.PARAMETER WhatIf
    Show what would be restored without actually importing anything.

# -----------------------------------------------------------------------------
#  SETUP & USAGE
# -----------------------------------------------------------------------------
#
#  Prerequisites
#  -------------
#    1. PowerShell 7 (recommended):
#         winget install --id Microsoft.PowerShell --source winget
#         pwsh
#    2. Power BI module:
#         Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
#    3. Sign in to Power BI:
#         Connect-PowerBIServiceAccount
#
#  Usage
#  -----
#    Preview what would be restored (no changes made):
#         .\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000 -WhatIf
#
#    Restore into a specific workspace:
#         .\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000 -TargetWorkspaceId <guid>
#
#    Restore into My Workspace (omit -TargetWorkspaceId):
#         .\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000
#
#    Skip items that already exist instead of overwriting:
#         .\02-RestorePowerBI.ps1 -BackupPath .\... -TargetWorkspaceId <guid> -ConflictAction Ignore
#
#  Notes
#  -----
#    - You need EDIT access to the target workspace.
#    - You may see repeated MSAL warnings during sign-in
#      ("WARNING SetAuthorityUri ... defaulting to MsSts"). These are harmless
#      and can be safely IGNORED.
#    - Restoring everything into ONE target workspace flattens the original
#      per-workspace structure. To mirror the original layout, run this script
#      once per workspace subfolder with the matching -TargetWorkspaceId.
# -----------------------------------------------------------------------------
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$BackupPath,

    [Parameter()]
    [string]$TargetWorkspaceId = '',

    [Parameter()]
    [ValidateSet('Ignore', 'Abort', 'Overwrite', 'CreateOrOverwrite')]
    [string]$ConflictAction = 'CreateOrOverwrite'
)

# --- Variables -------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$restoreLog = [System.Collections.Generic.List[object]]::new()

# --- Validate backup folder ------------------------------------------------
if (-not (Test-Path -LiteralPath $BackupPath)) {
    throw "Backup folder not found: $BackupPath"
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

# --- Resolve target workspace ----------------------------------------------
$targetLabel = 'My Workspace'
if ($TargetWorkspaceId) {
    try {
        $tws = Get-PowerBIWorkspace -Id $TargetWorkspaceId
        $targetLabel = "$($tws.Name) [$TargetWorkspaceId]"
    }
    catch {
        # Non-fatal: the import will still try with the given Id.
        $targetLabel = "[$TargetWorkspaceId]"
        Write-Warning "Could not resolve target workspace name; proceeding with the supplied Id."
    }
}
Write-Host "Restoring into: $targetLabel" -ForegroundColor Cyan
Write-Host "Conflict action: $ConflictAction" -ForegroundColor Cyan
Write-Host ""

# --- Find restorable files -------------------------------------------------
$files = Get-ChildItem -LiteralPath $BackupPath -Recurse -File -Include '*.pbix', '*.rdl' |
    Sort-Object FullName

if (-not $files -or $files.Count -eq 0) {
    Write-Host "No .pbix or .rdl files found under: $BackupPath" -ForegroundColor Yellow
    Write-Host "(Dashboards and dataflows are saved as .json and cannot be restored by this script.)" -ForegroundColor Yellow
    return
}

Write-Host ("Found {0} restorable file(s)." -f $files.Count) -ForegroundColor Green
Write-Host ""

# --- Restore loop ----------------------------------------------------------
foreach ($f in $files) {
    $ext  = $f.Extension.TrimStart('.').ToLower()
    $type = if ($ext -eq 'rdl') { 'PaginatedReport' } else { 'Report' }

    # Derive a friendly report name: strip the trailing _<guid> the backup added.
    $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $name = $name -replace '_[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$', ''
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $f.BaseName }

    $sourceWs = Split-Path -Path $f.DirectoryName -Leaf   # original workspace folder name

    if ($PSCmdlet.ShouldProcess("$type '$name' -> $targetLabel", 'Import')) {
        try {
            $importParams = @{
                Path           = $f.FullName
                Name           = $name
                ConflictAction = $ConflictAction
                ErrorAction    = 'Stop'
            }
            if ($TargetWorkspaceId) { $importParams['WorkspaceId'] = $TargetWorkspaceId }

            New-PowerBIReport @importParams | Out-Null

            Write-Host ("  Restored: {0,-16} {1}" -f $type, $name) -ForegroundColor Green
            $restoreLog.Add([pscustomobject]@{
                SourceWorkspace = $sourceWs; Type = $type; Name = $name
                File = $f.Name; Status = 'Restored'; Detail = $targetLabel
            })
        }
        catch {
            # Common when a report is bound to an external/shared dataset, or the
            # name conflicts under -ConflictAction Abort.
            Write-Warning ("  Failed:   {0} '{1}' - {2}" -f $type, $name, $_.Exception.Message)
            $restoreLog.Add([pscustomobject]@{
                SourceWorkspace = $sourceWs; Type = $type; Name = $name
                File = $f.Name; Status = 'Failed'; Detail = $_.Exception.Message
            })
        }
    }
    else {
        # -WhatIf path
        Write-Host ("  Would restore: {0,-16} {1}  (from '{2}')" -f $type, $name, $sourceWs) -ForegroundColor DarkCyan
        $restoreLog.Add([pscustomobject]@{
            SourceWorkspace = $sourceWs; Type = $type; Name = $name
            File = $f.Name; Status = 'WhatIf'; Detail = "Would import to $targetLabel"
        })
    }
}

# --- Summary ---------------------------------------------------------------
Write-Host ""
Write-Host "Restore summary" -ForegroundColor Cyan
Write-Host "---------------" -ForegroundColor Cyan
$restoreLog |
    Group-Object Status |
    Sort-Object Name |
    ForEach-Object {
        $color = switch ($_.Name) {
            'Restored' { 'Green' }
            'WhatIf'   { 'DarkCyan' }
            default    { 'Red' }   # Failed
        }
        Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Count) -ForegroundColor $color
    }

# Write a restore log next to the backup.
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $BackupPath "restore-log_$stamp.csv"
$restoreLog | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Log: $logFile" -ForegroundColor Green

Write-Host ""
Write-Host "Reminder - not restored by this script:" -ForegroundColor DarkGray
Write-Host "  Dashboards (.json) and dataflows (.json) require manual recreation." -ForegroundColor DarkGray
Write-Host "  After restore, re-check credentials, gateway, permissions, RLS, and refresh." -ForegroundColor DarkGray
