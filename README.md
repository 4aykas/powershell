# Autodesk PowerShell Tools

A collection of PowerShell scripts for managing, cleaning, and backing up Autodesk products on Windows.

## Scripts Overview

| Script | Purpose |
|--------|---------|
| `revit-server-backup.ps1` | Exports all Revit Server models to real `.rvt` files in a dated Desktop folder |
| `autodesk-uninstaller.ps1` | Fully uninstalls all Autodesk products and wipes all leftovers |
| `autodesk-telemetry.ps1` | Registers a daily scheduled task to clear Autodesk telemetry & usage data |

---

## 🏗️ Revit Server Backup

### Why this script exists

Revit Server does **not** store models as regular `.rvt` files on disk.  
Each model is saved as a **folder** ending in `.rvt` containing binary chunks and metadata — not something you can open in Revit directly. A standard file copy gives you an unrestorable archive.

This script uses **`revitservertool.exe createLocalRVT`** — the official Autodesk CLI tool shipped with every Revit / Revit Server installation — to assemble a proper, openable `.rvt` file per model, mirroring the full server folder tree.

---

### How it works
revitservertool.exe createLocalRVT "FolderName/ModelName.rvt"
-s SERVERHOSTNAME
-d "C:...\FolderName\ModelName.rvt"
-o
---

| Argument | Meaning |
|----------|---------|
| `createLocalRVT` | Command to assemble a real `.rvt` from server chunks |
| `"FolderName/Model.rvt"` | RSN model path (relative to Projects root) |
| `-s SERVERHOSTNAME` | Revit Server hostname (read from `RSN.ini`) |
| `-d "C:\...\Model.rvt"` | Full destination file path (mirrors server tree) |
| `-o` | Overwrite if file already exists |

---

### Script steps

1. **Auto-detect** — Scans all 4 known install path patterns for every version 2020–2027, shows `[TOOL OK]` / `[PROJECTS OK]` tags
2. **Version select** — Auto-selects if only one version found; shows a numbered menu if multiple
3. **Validate** — Confirms `revitservertool.exe` exists and shows its file version
4. **Read hostname** — Parses `RSN.ini` to get the server hostname automatically
5. **Scan models** — Walks the Projects folder and builds the full model list with relative paths
6. **Create backup folder** — Creates the destination on the Desktop:
7. **Export loop** — Calls `createLocalRVT` per model with `[1/N]` progress counter
8. **Skip locked** — Exit code `5` (locked) and `1` (busy) are caught silently; script continues
9. **Verify** — Checks each exported `.rvt` exists and logs its size
10. **Manifest** — Writes `_BACKUP_MANIFEST.txt` with `SUCCESS` / `FAILED` / `SKIPPED` per model

---

### Requirements & notes

- ✅ Run **on the Revit Server host machine** (tool and Projects folder already present)
- ✅ Supported versions: **Revit Server 2020 – 2027**
- ✅ Tool version must **match** the hosted server version
- ✅ No full Revit installation needed — `RevitServerToolCommand\` folder is self-contained
- ⚠️ Best practice: run **after hours** when no users are inside models
- ⚠️ Locked/busy models are **skipped**, not failed — check the manifest after each run

---

### Quick Run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/4aykas/powershell/main/revit-server-backup.ps1 | iex`""
```

---

## 🧹 Autodesk Uninstaller

Fully removes **all Autodesk products** from a Windows machine — MSI packages, registry keys, leftover program folders, ProgramData, and scheduled tasks.

### What it removes

- All Autodesk MSI/EXE-installed products (detected via registry)
- `C:\Program Files\Autodesk\`
- `C:\ProgramData\Autodesk\`
- `%APPDATA%\Autodesk\` and `%LOCALAPPDATA%\Autodesk\`
- Autodesk-related registry keys under `HKLM` and `HKCU`
- Autodesk scheduled tasks

### Requirements

- Windows 10 / 11
- PowerShell 5.1+
- **Administrator rights required**

### Quick Run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-uninstaller.ps1 | iex`""
```

---

## 📊 Autodesk Telemetry Cleaner

Registers a **daily Windows Scheduled Task** that clears Autodesk telemetry, usage analytics, and diagnostic data files automatically.

### What it clears

- Autodesk Analytics / telemetry JSON and log files
- AdSSO and identity service caches
- CER (Customer Error Reporting) data
- Usage data under `%APPDATA%` and `%LOCALAPPDATA%`

### Requirements

- Windows 10 / 11
- PowerShell 5.1+
- **Administrator rights required** (to register the scheduled task)

### Quick Run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-telemetry.ps1 | iex`""
```

---

## System Requirements

| Requirement | Details |
|-------------|---------|
| OS | Windows 10 / 11 |
| PowerShell | 5.1 or newer |
| Admin rights | Required for uninstaller, telemetry scheduler, and backup |
| Network access | Read access to Revit Server Projects share (backup only) |

---

## License

MIT — free to use, modify, and distribute.
