# Autodesk PowerShell Tools

A collection of PowerShell scripts for managing, cleaning, and maintaining Autodesk products on Windows.

> All scripts require **PowerShell 5.1+** and must be run as **Administrator**.

---

## Scripts

| Script | Purpose |
|---|---|
| [`autodesk-uninstaller.ps1`](./autodesk-uninstaller.ps1) | Fully uninstalls all Autodesk products and wipes every leftover |
| [`autodesk-telemetry.ps1`](./autodesk-telemetry.ps1) | Registers a daily scheduled task to auto-clear Autodesk telemetry and usage data |
| [`autodesk-trial.ps1`](./autodesk-trial.ps1) | Resets Autodesk trial license counters and clears related caches |

> **Revit Server Backup** has been moved to its own dedicated repository.

---

## Autodesk Uninstaller

Completely removes **all Autodesk products** from a Windows machine across 7 phases:

| Phase | What it does |
|---|---|
| 1 | Uninstalls all detected Autodesk products — MSI and ODIS-based (4 passes) |
| 2 | Stops and deletes Autodesk background services (FLEXnet, AdskLicensing, Desktop App) |
| 3 | Deletes leftover folders under `Program Files`, `ProgramData`, and all user profiles |
| 4 | Wipes FLEXnet and AdskLicensing data, clears `ADSKFLEX_LICENSE_FILE` env var |
| 5 | Removes telemetry, usage statistics, CER crash data, and ODIS metadata |
| 6 | Cleans Autodesk registry hives from both `HKLM` and `HKCU` (auto-backup to Desktop first) |
| 7 | Clears `%TEMP%` and `C:\Windows\Temp` |

### Products covered

- AutoCAD, Revit, Civil 3D, Navisworks, Inventor, Vault
- Maya, 3ds Max, ReCap, InfraWorks, Fusion, BIM 360
- Autodesk Access, Identity Manager, Genuine Service, Desktop Licensing

### Quick run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command \`"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-uninstaller.ps1 | iex\`""
```

> A registry backup is automatically saved to the Desktop before Phase 6 runs.  
> Restart the machine after completion before reinstalling any Autodesk product.

---

## Autodesk Telemetry Cleaner

Registers a **daily Windows Scheduled Task** that silently clears Autodesk telemetry, analytics, and diagnostic data in the background.

### What it clears

- Analytics and telemetry JSON/log files (`%LOCALAPPDATA%\Autodesk\Analytics`)
- AdSSO and identity service caches
- CER (Customer Error Reporting) crash data
- ADLM usage data under `%APPDATA%` and `%LOCALAPPDATA%`
- Autodesk Desktop App activity logs

### Quick run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command \`"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-telemetry.ps1 | iex\`""
```

---

## Autodesk Trial Reset

Resets Autodesk trial license counters and clears related registry entries and cache files, allowing a fresh trial period to begin.

### What it resets

- Trial counter registry entries under `HKCU` and `HKLM`
- AdLM product information pit files
- Identity and SSO token caches

### Quick run

```powershell
Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command \`"irm https://raw.githubusercontent.com/4aykas/powershell/main/autodesk-trial.ps1 | iex\`""
```

---

## System Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1 or newer |
| Privileges | Administrator required for all scripts |
| Network | Not required (all scripts run locally) |

---

## License

MIT — free to use, modify, and distribute.
