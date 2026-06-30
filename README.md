# Power BI & Fabric Inventory / Backup Scripts

> ⚠️ **EXPERIMENTAL.** These scripts are provided as-is, for convenience. They
> have not been exhaustively tested across every tenant configuration. Always run
> inventory and `-WhatIf` first, verify results, and test on a non-production
> workspace before relying on them. Treat any backup as best-effort, not as a
> guaranteed or complete recovery point. Use at your own risk.

A small set of PowerShell scripts to **inventory**, **back up**, and **restore**
content in a Microsoft Power BI / Fabric tenant.

These can be helpful for:

- **Backing up** Power BI content (reports + semantic models) to `.pbix` files
  you keep outside the service — handy for version history or before risky
  changes.
- **Migrating** content between workspaces or tenants: inventory the source,
  back up the `.pbix` / `.rdl` files, then restore them into the target
  workspace (review the migration caveats below — credentials, gateways, RLS,
  and external-dataset bindings still need manual attention).
- **Auditing / documenting** what exists across Power BI and Fabric workspaces.

| Script | Purpose |
|--------|---------|
| `01-InventoryPowerBI.ps1` | List all Power BI items, and optionally back them up to disk. |
| `02-RestorePowerBI.ps1` | Restore `.pbix` / `.rdl` files from a backup into a workspace. |
| `10-InventoryFabric.ps1` | List all Microsoft Fabric items across workspaces. |

> **Scope:** Power BI / Fabric content lives in a **tenant**, organized into
> **workspaces** — *not* in an Azure subscription. All scripts iterate workspaces.

---

## Prerequisites (all scripts)

1. **PowerShell 7** (recommended; installs alongside Windows PowerShell):
   ```powershell
   winget install --id Microsoft.PowerShell --source winget
   pwsh
   ```
2. **Power BI management module:**
   ```powershell
   Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
   ```
3. **Sign in once** (the same session is reused by all three scripts, including
   the Fabric one):
   ```powershell
   Connect-PowerBIServiceAccount
   ```

> **MSAL warnings:** During sign-in you may see repeated lines like
> `WARNING SetAuthorityUri ... defaulting to MsSts`. These are harmless
> diagnostic messages from the auth library and can be safely **ignored**.

> **Admin vs. user scope:** Where a script can use *Organization* (admin) scope,
> a Fabric administrator can enumerate **all** workspaces in the tenant — even
> ones they aren't a member of (read/enumerate visibility, not edit rights).
> Without admin rights, the scripts fall back to **user scope** and list only
> workspaces your account belongs to.

---

## 1. `01-InventoryPowerBI.ps1` — Power BI inventory & backup

Lists every classic Power BI item the `MicrosoftPowerBIMgmt` module can reach:

- **Reports** (standard) and **Paginated reports** (shown as separate types)
- **Semantic models** (datasets)
- **Dashboards**
- **Dataflows**

With `-Backup`, it also downloads what can be exported, into a subfolder per
workspace, and writes a log plus a `RESTORE-NOTES.txt` describing what is and
isn't restorable.

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-CsvPath <path>` | Also export the inventory table to a CSV file. |
| `-IncludePersonalWorkspace` | Include the signed-in user's own *My Workspace*. |
| `-AllPersonalWorkspaces` | Include **every** user's personal workspace (requires Fabric admin). |
| `-Backup` | Download `.pbix` / `.rdl` / dataflow `.json` for every exportable item. |
| `-BackupPath <path>` | Root folder for `-Backup` output. Default: `.\PowerBIBackup_<timestamp>`. |

### Examples

```powershell
# Inventory only, all accessible workspaces
.\01-InventoryPowerBI.ps1

# Inventory + CSV export
.\01-InventoryPowerBI.ps1 -CsvPath .\pbi-inventory.csv

# Include your personal My Workspace
.\01-InventoryPowerBI.ps1 -IncludePersonalWorkspace

# Back up everything possible (auto-named folder)
.\01-InventoryPowerBI.ps1 -Backup

# Back up to a specific folder
.\01-InventoryPowerBI.ps1 -Backup -BackupPath D:\PBIBackup
```

### What the backup can and cannot restore

| Item | Backup format | Restorable? |
|------|---------------|-------------|
| Report | `.pbix` (includes its semantic model) | **Yes** |
| Paginated report | `.rdl` | **Yes** |
| Dataflow | `.json` (model definition) | Partial — manual recreation, rebind credentials/gateway |
| Dashboard | `.json` (metadata + tiles) | **No** — there is no dashboard import API; reference only |

**Not captured by any export** (must be reconfigured after a restore): data
source credentials & gateway bindings, sharing/permissions, Row-Level Security
(RLS), scheduled refresh, app packaging, and workspace settings.

> Reports built on a **live connection / DirectQuery to a shared dataset**, and
> **system reports** (e.g. usage metrics), cannot be exported as `.pbix` and are
> skipped. Every skip/failure is recorded in `backup-log.csv`.

---

## 2. `02-RestorePowerBI.ps1` — Power BI restore

Scans a backup folder created by script 01, finds every `.pbix` and `.rdl`, and
re-imports each into a target workspace via `New-PowerBIReport`.

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-BackupPath <path>` | **Required.** Backup folder to restore from (searched recursively). |
| `-TargetWorkspaceId <guid>` | Workspace to restore into. Omit to restore into *My Workspace*. |
| `-ConflictAction <action>` | `Ignore` / `Abort` / `Overwrite` / `CreateOrOverwrite`. Default: `CreateOrOverwrite`. |
| `-WhatIf` | Preview what would be restored without importing anything. |

### Examples

```powershell
# Preview first (no changes made) - recommended
.\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000 -WhatIf

# Restore into a specific workspace
.\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000 -TargetWorkspaceId <guid>

# Restore into My Workspace
.\02-RestorePowerBI.ps1 -BackupPath .\PowerBIBackup_20260630_120000

# Skip items that already exist instead of overwriting
.\02-RestorePowerBI.ps1 -BackupPath .\... -TargetWorkspaceId <guid> -ConflictAction Ignore
```

### Notes & limitations

- You need **edit access** to the target workspace.
- **Dashboards and dataflows are not restored** (their `.json` files can't be
  imported); the script skips them and reminds you at the end.
- Reports bound to an **external/shared dataset** may fail to import (Power BI
  treats the embedded dataset reference as a conflict). Such failures are caught
  and logged — you'd rebind them to the target dataset afterward.
- Restoring everything into **one** workspace flattens the original per-workspace
  layout. To mirror the original structure, run the script once per workspace
  subfolder with the matching `-TargetWorkspaceId`.
- A timestamped `restore-log_<stamp>.csv` is written into the backup folder.

---

## 3. `10-InventoryFabric.ps1` — Microsoft Fabric inventory

Lists every **Fabric** item across workspaces via the Fabric REST API
(`https://api.fabric.microsoft.com/v1`), reusing your Power BI sign-in. Inventory
only — it does **not** back anything up.

Covers all item types the Fabric API returns, e.g. Lakehouses, Warehouses,
Notebooks, Data pipelines, Eventhouses, KQL databases, ML models, Semantic
models, Reports, and more. Handles API pagination automatically.

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-CsvPath <path>` | Also export the inventory table to a CSV file. |
| `-WorkspaceId <guid[,guid...]>` | Scan only the given workspace(s). Omit to scan all accessible workspaces. |

### Examples

```powershell
# All accessible workspaces
.\10-InventoryFabric.ps1

# Export to CSV
.\10-InventoryFabric.ps1 -CsvPath .\fabric-inventory.csv

# Specific workspaces only
.\10-InventoryFabric.ps1 -WorkspaceId 1111-...,2222-...
```

### Power BI vs. Fabric — why the lists may match

Fabric and Power BI are two views of the **same** underlying objects: a Power BI
report is a Fabric item of type `Report`, a semantic model is type
`SemanticModel`, and so on.

- If your tenant contains **only classic Power BI content**, both inventory
  scripts will list largely the **same items** — this is expected, not a bug.
- You'll see Fabric list **more** only when workspaces are on a **Fabric capacity**
  (an F-SKU, or Premium/Trial with Fabric enabled) and actually contain
  Fabric-native items (Lakehouse, Warehouse, Notebook, etc.).
- A few items differ either way: **dashboards** and **Dataflow Gen1** appear in
  the Power BI script but not as Fabric items; **Dataflow Gen2** appears as a
  Fabric item.

---

## Suggested workflow

```powershell
# 0. Sign in once
Connect-PowerBIServiceAccount

# 1. See what you have (Power BI)
.\01-InventoryPowerBI.ps1 -CsvPath .\pbi-inventory.csv

# 2. See what you have (Fabric)
.\10-InventoryFabric.ps1 -CsvPath .\fabric-inventory.csv

# 3. Back up Power BI content
.\01-InventoryPowerBI.ps1 -Backup -BackupPath D:\PBIBackup

# 4. (If ever needed) restore - preview first, then run for real
.\02-RestorePowerBI.ps1 -BackupPath D:\PBIBackup\PowerBIBackup_<stamp> -WhatIf
.\02-RestorePowerBI.ps1 -BackupPath D:\PBIBackup\PowerBIBackup_<stamp> -TargetWorkspaceId <guid>
```

---

## Important caveats

- **Experimental / best-effort.** Not exhaustively tested across tenant
  configurations. Validate results and test on non-production workspaces first.
- This is a **content** backup (report definitions + models), **not** a full
  disaster-recovery snapshot of the tenant. Permissions, credentials, gateways,
  RLS, refresh schedules, and dashboards are not part of a restorable backup.
- For **migrations**, the inventory + backup + restore flow moves the report and
  model definitions, but the target still needs manual setup: re-enter data
  source credentials, rebind gateways, reconfigure sharing/RLS/refresh, and
  rebind any reports that point to an external/shared dataset.
- **Lakehouse / Warehouse data** is *not* exported by any of these scripts. The
  Fabric script lists those items but their underlying data (OneLake / Delta)
  would need separate tooling (e.g. `azcopy` / ADLS tools).
- The scripts read the most they can with your current permissions; workspaces or
  items you can't access are reported and skipped rather than failing the run.
