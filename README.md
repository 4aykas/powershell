# Autodesk PowerShell Tools

A collection of PowerShell scripts for managing, cleaning, and backing up Autodesk products on Windows.

## Scripts

| Script | Purpose |
|--------|---------|
| `revit-server-backup.ps1` | Exports all Revit Server models to real `.rvt` files in a dated Desktop folder |
| `autodesk-uninstaller.ps1` | Fully uninstalls all Autodesk products and wipes all leftovers |
| `autodesk-telemetry.ps1` | Registers a daily scheduled task to clear Autodesk telemetry & usage data |

---

> Run directly without downloading — right-click PowerShell → **Run as Administrator**, then paste the one-liner.

---

## 🏗️ Revit Server Backup

### Why this script exists

Revit Server does **not** store models as regular `.rvt` files on disk. Each model is saved as a **folder** ending in `.rvt` containing binary chunks and metadata — not something you can open in Revit directly. A standard file copy gives you an unrestorable mess.

This script uses `revitservertool.exe createLocalRVT` — the **official Autodesk CLI tool** shipped with every Revit / Revit Server installation — to assemble a proper, openable `.rvt` file per model, mirroring the full server folder tree.

### What it does (10 steps)

1. Auto-detects all installed Revit versions (2020–2027) and their `revitservertool.exe` paths
2. Shows a version selection menu (auto-selects if only one found)
3. Validates the tool path and the Projects folder
4. Reads `RSN.ini` to get the server hostname
5. Scans the Projects folder and builds the full model list
6. Creates the backup destination on the Desktop:
