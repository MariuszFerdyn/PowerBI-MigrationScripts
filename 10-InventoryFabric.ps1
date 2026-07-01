<#
.SYNOPSIS
    Enumerates all Microsoft Fabric items across workspaces in the tenant
    (Lakehouses, Warehouses, Notebooks, Pipelines, Semantic models, Reports,
    Eventhouses, KQL databases, ML models, and every other item type).

.DESCRIPTION
    Calls the Fabric REST API (https://api.fabric.microsoft.com/v1) using the
    token from your Power BI sign-in, lists every workspace, then lists every
    item in each workspace. Output is grouped by workspace with a per-type
    summary. Handles API pagination (continuationToken).

    This is INVENTORY ONLY - it does not download or back up anything.

    NOTE: Fabric content is scoped to a TENANT (not an Azure subscription).
    Items live in workspaces, so this script iterates workspaces.

    NOTE: This lists items the Fabric REST API exposes. It complements
    01-InventoryPowerBI.ps1, which uses the older Power BI cmdlets. Fabric items
    (Lakehouse, Warehouse, Notebook, Pipeline, etc.) appear here; classic
    Power BI specifics (dashboards, dataflow Gen1) appear in the Power BI script.

.PARAMETER CsvPath
    Optional path to export the results as a CSV file.

.PARAMETER WorkspaceId
    One or more specific workspace Ids to scan. If omitted, all accessible
    workspaces are scanned.

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
#    2. Install the Power BI module (provides the authenticated REST helper):
#         Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
#
#    3. Sign in (this token also authorizes Fabric API calls):
#         Connect-PowerBIServiceAccount
#
#  Usage
#  -----
#    Scan all accessible workspaces:
#         .\10-InventoryFabric.ps1
#
#    Export results to CSV:
#         .\10-InventoryFabric.ps1 -CsvPath .\fabric-inventory.csv
#
#    Scan specific workspaces only:
#         .\10-InventoryFabric.ps1 -WorkspaceId 1111-...,2222-...
#
#  Notes
#  -----
#    - Uses the Fabric REST API via Invoke-PowerBIRestMethod, so the same
#      Connect-PowerBIServiceAccount session is reused.
#    - Personal "My workspace" (type Personal) is included automatically when
#      the API returns it for your account.
#    - You may see repeated MSAL warnings during sign-in
#      ("WARNING SetAuthorityUri ... defaulting to MsSts"). These are harmless
#      and can be safely IGNORED.
#    - Listing items requires Viewer (or higher) on each workspace. Workspaces
#      you cannot read are reported and skipped.
# -----------------------------------------------------------------------------
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvPath = '',

    [Parameter()]
    [string[]]$WorkspaceId
)

# --- Variables -------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$FabricBase = 'https://api.fabric.microsoft.com/v1'
$results    = [System.Collections.Generic.List[object]]::new()

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

# Emits a "could not read" warning. If the error is Unauthorized (401/403), it
# appends a hint: that status is USUALLY harmless (a workspace you can enumerate
# but not read, e.g. someone else's personal workspace) but CAN signal a real
# problem (expired/insufficient token, revoked admin role, Conditional Access/MFA
# block, or a workspace not on a Fabric capacity you can access).
function Write-ReadWarning {
    param(
        [string]$What,        # e.g. 'items'
        [string]$Workspace,
        [string]$Message
    )
    $line = "Could not read $What in '$Workspace': $Message"
    if ($Message -match 'Unauthorized|401|403') {
        $line += " [Unauthorized - usually just a workspace you can't read (e.g. a personal workspace); "
        $line += "but if this appears for workspaces you SHOULD be able to read, or for all of them, "
        $line += "check your sign-in/admin role and workspace access.]"
    }
    Write-Warning $line
}

# --- Helper: GET a Fabric API URL, following continuationToken pagination ----
function Invoke-FabricGetAll {
    param(
        [Parameter(Mandatory)] [string]$Url,   # absolute Fabric API URL
        [Parameter(Mandatory)] [string]$ValueProperty  # 'value' (core) name of the array property
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Url
    while ($next) {
        $raw  = Invoke-PowerBIRestMethod -Url $next -Method Get
        $data = $raw | ConvertFrom-Json
        if ($data.$ValueProperty) { $items.AddRange(@($data.$ValueProperty)) }

        if ($data.continuationUri) {
            $next = $data.continuationUri
        }
        elseif ($data.continuationToken) {
            $sep  = if ($Url -match '\?') { '&' } else { '?' }
            $next = "{0}{1}continuationToken={2}" -f $Url, $sep, [uri]::EscapeDataString($data.continuationToken)
        }
        else {
            $next = $null
        }
    }
    return $items
}

# --- Get workspaces ---------------------------------------------------------
$workspaces = [System.Collections.Generic.List[object]]::new()

if ($WorkspaceId -and $WorkspaceId.Count -gt 0) {
    foreach ($id in $WorkspaceId) {
        try {
            $ws = Invoke-PowerBIRestMethod -Url "$FabricBase/workspaces/$id" -Method Get | ConvertFrom-Json
            $workspaces.Add($ws)
        }
        catch {
            Write-ReadWarning -What 'workspace' -Workspace $id -Message $_.Exception.Message
        }
    }
}
else {
    Write-Host "Listing all accessible Fabric workspaces..." -ForegroundColor Cyan
    try {
        $all = Invoke-FabricGetAll -Url "$FabricBase/workspaces" -ValueProperty 'value'
        $workspaces.AddRange(@($all))
    }
    catch {
        throw "Failed to list Fabric workspaces: $($_.Exception.Message)"
    }
}

Write-Host ("Found {0} workspace(s)." -f $workspaces.Count) -ForegroundColor Green
Write-Host ""

# --- Iterate ----------------------------------------------------------------
foreach ($ws in $workspaces) {
    $wsName = $ws.displayName
    $wsId   = $ws.id
    Write-Host "Scanning workspace: $wsName" -ForegroundColor Cyan

    try {
        $items = Invoke-FabricGetAll -Url "$FabricBase/workspaces/$wsId/items" -ValueProperty 'value'
        foreach ($it in $items) {
            $results.Add([pscustomobject]@{
                Workspace     = $wsName
                WorkspaceId   = $wsId
                Type          = $it.type
                Name          = $it.displayName
                Id            = $it.id
                Description   = $it.description
            })
        }
    }
    catch {
        Write-ReadWarning -What 'items' -Workspace "$wsName ($wsId)" -Message $_.Exception.Message
    }
}

# --- Output -----------------------------------------------------------------
if ($results.Count -eq 0) {
    Write-Host ""
    Write-Host "No Fabric items found in the scanned workspaces." -ForegroundColor Yellow
    Write-Host "(If you expected results, confirm you have access and that the workspaces are on a Fabric capacity.)" -ForegroundColor Yellow
    return
}

$results |
    Sort-Object Workspace, Type, Name |
    Format-Table Workspace, Type, Name, Id -AutoSize

Write-Host ""
Write-Host ("Total items: {0}  across {1} workspace(s)" -f $results.Count, $workspaces.Count) -ForegroundColor Green

# Per-type summary (dynamic - Fabric has many item types and adds more over time).
Write-Host "By type:" -ForegroundColor Green
$results |
    Group-Object Type |
    Sort-Object Count -Descending |
    ForEach-Object { Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count) -ForegroundColor Green }

if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
    $results | Sort-Object Workspace, Type, Name | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Exported to: $CsvPath" -ForegroundColor Green
}
