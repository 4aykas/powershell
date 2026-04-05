
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
