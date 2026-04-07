# Autodesk PowerShell Tools

A collection of PowerShell scripts for managing, cleaning, and maintaining Autodesk products on Windows.

> All scripts require **PowerShell 5.1+** and must be run as **Administrator**.

---

## Scripts

| Script | Purpose |
|---|---|
| [`autodesk-uninstaller.ps1`](./autodesk-uninstaller.ps1) | Fully uninstalls all Autodesk products and wipes every leftover (7 phases) |
| [`autodesk-telemetry.ps1`](./autodesk-telemetry.ps1) | Registers a daily scheduled task that silently clears Autodesk telemetry, logs, and usage data |

> **Revit Server Backup** has been moved to its own dedicated repository.

---

## autodesk-uninstaller.ps1

Completely removes **all Autodesk products** from a Windows machine in 7 sequential phases.

### Phases

| Phase | What it does |
|---|---|
| 1 | Detects and uninstalls all Autodesk products — MSI and ODIS-based (4 passes) |
| 2 | Stops and deletes Autodesk background services (FLEXnet, AdskLicensing, Desktop App) |
| 3 | Deletes leftover folders under `Program Files`, `ProgramData`, and all user profiles |
| 4 | Wipes FLEXnet / AdskLicensing data and clears the `ADSKFLEX_LICENSE_FILE` environment variable |
| 5 | Removes telemetry, usage statistics, CER crash reports, and ODIS installer metadata |
| 6 | Cleans Autodesk registry hives from `HKLM` and `HKCU` (auto-backup to Desktop first) |
| 7 | Clears `%TEMP%` and `C:\Windows\Temp` |

### Products covered

AutoCAD, Revit, Civil 3D, Navisworks, Inventor, Vault, Maya, 3ds Max, ReCap, InfraWorks, Fusion 360, BIM 360, Autodesk Access, Identity Manager, Genuine Service, Desktop Licensing.

### Configuration flags

Edit the top of the script to control behaviour:

```powershell
$DeleteCAutodeskFolder = $true   # Delete C:\Autodesk install cache (can be 10+ GB)
$DeleteUserProfiles    = $true   # Delete Autodesk folders from all user profiles
$CleanRegistry         = $true   # Remove Autodesk registry hives
$CleanFLEXnet          = $true   # Wipe FLEXnet / AdskLicensing data
$CleanTelemetry        = $true   # Remove telemetry, logs, and usage statistics
$AutoRestart           = $false  # Automatically restart when done
```

### Run

```powershell
Start-Process powershell -Verb RunAs `
  -ArgumentList "-ExecutionPolicy Bypass -File autodesk-uninstaller.ps1"
```

> A registry export is automatically saved to the Desktop before Phase 6 runs.  
> Restart the machine after completion before reinstalling any Autodesk product.

---

## autodesk-telemetry.ps1

Registers a **daily Windows Scheduled Task** that runs at 03:00 as SYSTEM and silently clears all Autodesk telemetry, analytics, and diagnostic data in the background. Run this script **once** as Administrator — the task then runs automatically every day.

### What gets cleared daily

- Analytics and telemetry data (`%LOCALAPPDATA%\Autodesk\Analytics`)
- ADLM usage data (`%APPDATA%\Autodesk\ADLM`, `%LOCALAPPDATA%\Autodesk\ADLM`)
- CER (Customer Error Reporting) crash data
- Autodesk Desktop App logs and cache
- AdskLicensing logs and cache
- Revit journal files and collaboration cache (versions 2020–2027)
- AutoCAD, Civil 3D, Navisworks, Inventor, Maya, 3ds Max, Vault, Fusion 360 logs and cache (versions 2020–2027)
- Wildcard sweep for `*.adsklog`, `Adsk*`, `autodesk*` in `%TEMP%` and `C:\Windows\Temp`
- Completion timestamp written to `C:\ProgramData\Autodesk\CleanupLog\cleanup.log` (last 30 entries kept)

### Task details

| Setting | Value |
|---|---|
| Task name | `Autodesk Daily Telemetry Cleanup` |
| Runs as | `SYSTEM` |
| Schedule | Daily at `03:00` (configurable at top of script) |
| Execution limit | 1 hour |
| Start when available | Yes (catches missed runs) |

### Run

```powershell
Start-Process powershell -Verb RunAs `
  -ArgumentList "-ExecutionPolicy Bypass -File autodesk-telemetry.ps1"
```

### Useful task commands

```powershell
# Run immediately
Start-ScheduledTask -TaskName "Autodesk Daily Telemetry Cleanup"

# Remove task
Unregister-ScheduledTask -TaskName "Autodesk Daily Telemetry Cleanup"

# Check last run
Get-ScheduledTaskInfo -TaskName "Autodesk Daily Telemetry Cleanup"

# View cleanup log
Get-Content "C:\ProgramData\Autodesk\CleanupLog\cleanup.log"
```

---

## System Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1 or newer |
| Privileges | Administrator required for all scripts |
| Network | Not required — all scripts run locally |

---

## License

MIT — free to use, modify, and distribute.
