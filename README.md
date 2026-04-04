# 🧹 Autodesk PowerShell Toolkit

A collection of PowerShell scripts for managing, cleaning, and backing up Autodesk products on Windows.

---

## 📦 Scripts

| Script | Purpose |
|---|---|
| `autodesk-uninstaller.ps1` | Fully uninstalls all Autodesk products and wipes all leftovers |
| `autodesk-telemetry.ps1` | Registers a daily scheduled task to clear Autodesk telemetry & usage data |
| `revit-server-backup.ps1` | Backs up a Revit Server Projects tree to a dated folder on the Desktop |

---

## ⚡ Quick Run (from URL)

> Run directly without downloading — **requires PowerShell as Administrator**

```powershell
# Uninstaller
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-uninstaller.ps1 | iex`""

# Telemetry cleanup scheduler
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-telemetry.ps1 | iex`""

# Revit Server backup
irm https://raw.githubusercontent.com/4aykas/powershell/main/revit-server-backup.ps1 | iex
```

---

## 🗑️ autodesk-uninstaller.ps1

Uninstalls **all** Autodesk products and removes every trace — files, folders, registry, FLEXnet licenses, telemetry, and temp files.

```powershell
.\autodesk-uninstaller.ps1
```

**What it removes:**
- All installed Autodesk products (MSI + ODIS, up to 4 passes)
- Background services (AdskLicensing, FLEXnet, Desktop App)
- `C:\Program Files\Autodesk`, `C:\ProgramData\Autodesk`, `C:\Autodesk`
- `%AppData%\Autodesk`, `%LocalAppData%\Autodesk` (all user profiles)
- `C:\ProgramData\FLEXnet` — license trusted storage
- Registry: `HKLM\SOFTWARE\Autodesk`, `HKCU\SOFTWARE\Autodesk`, FLEXlm keys
- All temp files, logs, and stale uninstall entries

> ⚠️ Auto-exports a registry backup to Desktop before deleting. **Restart required after.**

---

## 📅 autodesk-telemetry.ps1

Registers a **daily Windows Scheduled Task** (runs at 03:00 as SYSTEM) that silently wipes Autodesk telemetry, usage stats, analytics, and logs.

```powershell
# Register the task (run ONCE as Administrator)
.\autodesk-telemetry.ps1

# Run immediately after registering
Start-ScheduledTask -TaskName "Autodesk Daily Telemetry Cleanup"

# Remove the task
Unregister-ScheduledTask -TaskName "Autodesk Daily Telemetry Cleanup"
```

**Cleans daily (versions 2020–2027):**
- ADLM usage & license telemetry
- CER crash reports & analytics
- Revit Journal files, AutoCAD & Civil 3D cache
- FLEXnet license event logs, ODIS metadata
- `*.adsklog` temp files

**Log file:** `C:\ProgramData\Autodesk\CleanupLog\cleanup.log`

---

## 💾 revit-server-backup.ps1

Interactive backup — mirrors a full Revit Server Projects tree to a dated folder on the Desktop using Robocopy.

```powershell
# Run on the Revit Server host (or any machine with UNC access)
.\revit-server-backup.ps1
```

**Backup output structure:**
