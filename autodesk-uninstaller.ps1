#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Complete Autodesk Uninstaller & Cleanup Script
    Enhanced version with full leftover removal, registry cleanup,
    licensing data, usage statistics, and telemetry wipeout.

.DESCRIPTION
    https://github.com/tv2vn/powershell/blob/main/autodesk-uninstaller.ps1
    Enhanced with recommendations from:
    - Autodesk Official Clean Uninstall Guide
    - Hagerman & Pentagon Solutions BIM Admin Best Practices
    - Community knowledge (forums.autodesk.com, Reddit r/AutoCAD, r/RevitForum)

.NOTES
    Run as Administrator in PowerShell.
    Restart the computer after completion.
    WARNING: This removes ALL Autodesk products and data.
#>

# ─────────────────────────────────────────────────────────────
# CONFIGURATION — Adjust as needed
# ─────────────────────────────────────────────────────────────
$DeleteCAutodeskFolder  = $true   # Delete C:\Autodesk (install cache, can be 10+ GB)
$DeleteUserProfiles     = $true   # Delete per-user AppData Autodesk folders (all users)
$CleanRegistry          = $true   # Delete Autodesk registry hives
$CleanFLEXnet           = $true   # Delete FLEXnet / AdskLicensing data
$CleanTelemetry         = $true   # Delete usage stats, telemetry, log files
$AutoRestart            = $false  # Set to $true to auto-restart when done
# ─────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Message, [string]$Color = "Cyan")
    Write-Host "`n[$([datetime]::Now.ToString('HH:mm:ss'))] $Message" -ForegroundColor $Color
}

function Remove-PathSilently {
    param([string]$Path)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Host "  Deleted: $Path" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Skipped (in use): $Path" -ForegroundColor DarkYellow
        }
    }
}

function Remove-RegistryKeySilently {
    param([string]$KeyPath)
    if (Test-Path $KeyPath) {
        try {
            Remove-Item -Path $KeyPath -Recurse -Force -ErrorAction Stop
            Write-Host "  Deleted registry key: $KeyPath" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Could not delete: $KeyPath — $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}

# ─────────────────────────────────────────────────────────────
# PHASE 1 — Uninstall all Autodesk products (runs 4x to
#            catch dependent components removed in later passes)
# ─────────────────────────────────────────────────────────────
function Invoke-AutodeskUninstaller {
    Clear-Host
    Write-Step "Scanning registry for installed Autodesk products..." "Yellow"

    $apps = @()
    $apps  = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
    $apps += Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
    $apps  = $apps | Where-Object {
        ($_.DisplayName -like "*Autodesk*")    -or
        ($_.Publisher   -like "*Autodesk*")    -or
        ($_.DisplayName -like "*AutoCAD*")     -or
        ($_.DisplayName -like "*Revit*")       -or
        ($_.DisplayName -like "*Maya*")        -or
        ($_.DisplayName -like "*3ds Max*")     -or
        ($_.DisplayName -like "*Navisworks*")  -or
        ($_.DisplayName -like "*Civil 3D*")    -or
        ($_.DisplayName -like "*BIM 360*")     -or
        ($_.DisplayName -like "*Dynamo*")      -or
        ($_.DisplayName -like "*ReCap*")       -or
        ($_.DisplayName -like "*InfraWorks*")  -or
        ($_.DisplayName -like "*Vault*")       -or
        ($_.DisplayName -like "*SketchBook*")  -or
        ($_.DisplayName -like "*Inventor*")    -or
        ($_.DisplayName -like "*Fusion*")      -or
        ($_.DisplayName -like "*Shotgun*")     -or
        ($_.DisplayName -like "*AdODIS*")      -or
        ($_.DisplayName -like "*AdskIdentity*")
    }
    $apps = $apps | Select-Object DisplayName, Publisher, PSChildName, UninstallString -Unique

    Write-Host "Found $($apps.Count) installed Autodesk product(s)" -ForegroundColor Yellow

    foreach ($app in $apps) {
        # ── Autodesk Access (ODIS-based) ──────────────────────────────
        if ($app.DisplayName -match "Autodesk Access") {
            Write-Step "Uninstalling Autodesk Access..." "Yellow"
            Start-Process -FilePath "C:\Program Files\Autodesk\AdODIS\V1\Installer.exe" `
                -ArgumentList "-q -i uninstall --trigger_point system -m C:\ProgramData\Autodesk\ODIS\metadata\{A3158B3E-5F28-358A-BF1A-9532D8EBC811}\pkg.access.xml -x `"C:\Program Files\Autodesk\AdODIS\V1\SetupRes\manifest.xsd`" --manifest_type package" `
                -NoNewWindow -Wait
        }

        # ── Autodesk Identity Manager ─────────────────────────────────
        if ($app.DisplayName -match "Autodesk Identity Manager") {
            Write-Step "Uninstalling Autodesk Identity Manager..." "Yellow"
            Start-Process -FilePath "C:\Program Files\Autodesk\AdskIdentityManager\uninstall.exe" `
                -ArgumentList "--mode unattended" -NoNewWindow -Wait
        }

        # ── Autodesk Genuine Service ──────────────────────────────────
        if ($app.DisplayName -match "Autodesk Genuine Service") {
            Write-Step "Uninstalling Autodesk Genuine Service..." "Yellow"
            Remove-PathSilently "$Env:ALLUSERSPROFILE\Autodesk\Adlm\ProductInformation.pit"
            Remove-PathSilently "$Env:USERPROFILE\AppData\Local\Autodesk\Genuine Autodesk Service\id.dat"
            Start-Process msiexec.exe -ArgumentList "/x `"{21DE6405-91DE-4A69-A8FB-483847F702C6}`" /qn" -NoNewWindow -Wait
        }

        # ── Autodesk Desktop Licensing Service ────────────────────────
        if ($app.DisplayName -match "Autodesk Desktop Licensing") {
            Write-Step "Uninstalling Autodesk Desktop Licensing Service..." "Yellow"
            $adskLicPath = "C:\Program Files (x86)\Common Files\Autodesk Shared\AdskLicensing\uninstall.exe"
            if (Test-Path $adskLicPath) {
                Start-Process -FilePath $adskLicPath -ArgumentList "--mode unattended" -NoNewWindow -Wait
            } else {
                Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn" -NoNewWindow -Wait
            }
        }

        # ── Carbon Insights for Revit ─────────────────────────────────
        if ($app.DisplayName -like "*Carbon Insights for Revit*") {
            Write-Step "Uninstalling Carbon Insights for Revit..." "Yellow"
            Start-Process -FilePath "C:\Program Files\Autodesk\AdODIS\V1\Installer.exe" `
                -ArgumentList "-q -i uninstall --trigger_point system -m C:\ProgramData\Autodesk\ODIS\metadata\{006E0C25-2C15-39A8-8590-AA5AD7D395D4}\pkg.RTCA.xml -x `"C:\Program Files\Autodesk\AdODIS\V1\SetupRes\manifest.xsd`" --manifest_type package" `
                -NoNewWindow -Wait
        }

        # ── ODIS installer-based product ──────────────────────────────
        if ($app.UninstallString -like "*installer.exe*") {
            Write-Step "Uninstalling $($app.DisplayName) (ODIS)..." "Yellow"
            Start-Process -FilePath "C:\Program Files\Autodesk\AdODIS\V1\Installer.exe" `
                -ArgumentList "-q -i uninstall --trigger_point system -m C:\ProgramData\Autodesk\ODIS\metadata\$($app.PSChildName)\bundleManifest.xml -x C:\ProgramData\Autodesk\ODIS\metadata\$($app.PSChildName)\SetupRes\manifest.xsd" `
                -NoNewWindow -Wait
            Start-Sleep -Seconds 3

        # ── MSI-based product ─────────────────────────────────────────
        } else {
            Write-Step "Uninstalling $($app.DisplayName) (MSI)..." "Yellow"
            Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn" -NoNewWindow -Wait
            Start-Sleep -Seconds 3
        }
    }
}

Write-Step "=== PHASE 1: Uninstalling Autodesk Products (up to 4 passes) ===" "Magenta"
for ($pass = 1; $pass -le 4; $pass++) {
    Write-Step "Pass $pass of 4..." "Cyan"
    Invoke-AutodeskUninstaller
}

# ─────────────────────────────────────────────────────────────
# PHASE 2 — Stop and remove Autodesk background services
# ─────────────────────────────────────────────────────────────
Write-Step "=== PHASE 2: Stopping Autodesk Background Services ===" "Magenta"

$autodeskServices = @(
    "AdskLicensingService", "AdskLicensingAgent",
    "Autodesk Desktop App Service", "AutodeskDesktopApp",
    "FNPLicensingService", "FlexNet Licensing Service",
    "FlexNet Licensing Service 64"
)

foreach ($svc in $autodeskServices) {
    $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host "  Stopping service: $svc" -ForegroundColor DarkGray
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
        Set-Service  -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue
        sc.exe delete $svc | Out-Null
    }
}

# ─────────────────────────────────────────────────────────────
# PHASE 3 — Delete leftover files and folders
# ─────────────────────────────────────────────────────────────
Write-Step "=== PHASE 3: Deleting Leftover Files & Folders ===" "Magenta"

$systemFolders = @(
    "C:\Program Files\Autodesk",
    "C:\Program Files\Common Files\Autodesk Shared",
    "C:\Program Files (x86)\Autodesk",
    "C:\Program Files (x86)\Common Files\Autodesk Shared",
    "C:\ProgramData\Autodesk",
    "C:\ProgramData\FLEXnet",
    "C:\Users\Public\Documents\Autodesk"
)

foreach ($folder in $systemFolders) { Remove-PathSilently $folder }
if ($DeleteCAutodeskFolder) { Remove-PathSilently "C:\Autodesk" }

$userFolders = @(
    "$Env:APPDATA\Autodesk", "$Env:LOCALAPPDATA\Autodesk",
    "$Env:LOCALAPPDATA\Autodesk Desktop App",
    "$Env:LOCALAPPDATA\Autodesk\Web Services",
    "$Env:LOCALAPPDATA\Autodesk\Genuine Autodesk Service",
    "$Env:LOCALAPPDATA\Autodesk\AdskLicensing"
)
foreach ($folder in $userFolders) { Remove-PathSilently $folder }

if ($DeleteUserProfiles) {
    Write-Step "Cleaning Autodesk folders from all user profiles..." "Cyan"
    Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        @(
            "$($_.FullName)\AppData\Roaming\Autodesk",
            "$($_.FullName)\AppData\Local\Autodesk",
            "$($_.FullName)\AppData\Local\Autodesk Desktop App",
            "$($_.FullName)\AppData\Local\Autodesk\AdskLicensing"
        ) | ForEach-Object { Remove-PathSilently $_ }
    }
}

# ─────────────────────────────────────────────────────────────
# PHASE 4 — FLEXnet & Licensing Data
# ─────────────────────────────────────────────────────────────
if ($CleanFLEXnet) {
    Write-Step "=== PHASE 4: Removing FLEXnet & Licensing Data ===" "Magenta"

    @(
        "C:\ProgramData\FLEXnet",
        "$Env:ALLUSERSPROFILE\Autodesk\Adlm",
        "$Env:LOCALAPPDATA\Autodesk\AdskLicensing",
        "C:\Program Files (x86)\Common Files\Autodesk Shared\AdskLicensing",
        "$Env:ALLUSERSPROFILE\Autodesk\Adlm\ProductInformation.pit"
    ) | ForEach-Object { Remove-PathSilently $_ }

    [System.Environment]::SetEnvironmentVariable("ADSKFLEX_LICENSE_FILE", $null, "Machine")
    [System.Environment]::SetEnvironmentVariable("ADSKFLEX_LICENSE_FILE", $null, "User")
    Write-Host "  Cleared ADSKFLEX_LICENSE_FILE environment variable" -ForegroundColor DarkGray
}

# ─────────────────────────────────────────────────────────────
# PHASE 5 — Telemetry, Usage Statistics & Log Files
# ─────────────────────────────────────────────────────────────
if ($CleanTelemetry) {
    Write-Step "=== PHASE 5: Removing Telemetry, Usage Stats & Logs ===" "Magenta"

    @(
        "$Env:APPDATA\Autodesk\ADLM",       # Usage-based licensing telemetry
        "$Env:LOCALAPPDATA\Autodesk\Analytics",
        "$Env:LOCALAPPDATA\Autodesk\Web Services",
        "$Env:ALLUSERSPROFILE\Autodesk\ADLM",
        "C:\ProgramData\Autodesk\ADLM",
        "$Env:LOCALAPPDATA\Autodesk Desktop App",
        "$Env:APPDATA\Autodesk Desktop App",
        "$Env:TEMP\Autodesk",
        "C:\Windows\Temp\Autodesk",
        "C:\ProgramData\Autodesk\ODIS",     # ODIS installer metadata
        "$Env:APPDATA\Autodesk\CER",        # Customer Error Reporting / crash data
        "$Env:LOCALAPPDATA\Autodesk\CER"
    ) | ForEach-Object { Remove-PathSilently $_ }

    Get-ChildItem -Path $Env:TEMP -Filter "*.adsklog" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path $Env:TEMP -Filter "Adsk*"     -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  Cleaned Autodesk temp/log files" -ForegroundColor DarkGray
}

# ─────────────────────────────────────────────────────────────
# PHASE 6 — Registry Cleanup
# ─────────────────────────────────────────────────────────────
if ($CleanRegistry) {
    Write-Step "=== PHASE 6: Cleaning Registry ===" "Magenta"

    # Auto-backup before touching registry
    $backupPath = "$Env:USERPROFILE\Desktop\Autodesk_RegistryBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    reg export "HKLM\SOFTWARE\Autodesk" $backupPath /y 2>$null
    Write-Host "  Registry backup saved to: $backupPath" -ForegroundColor Green

    @(
        "HKLM:\SOFTWARE\Autodesk",
        "HKLM:\SOFTWARE\WOW6432Node\Autodesk",
        "HKCU:\SOFTWARE\Autodesk",
        "HKCU:\SOFTWARE\FLEXlm License Manager",
        "HKLM:\SOFTWARE\FLEXlm License Manager",
        "HKLM:\SOFTWARE\WOW6432Node\FLEXlm License Manager",
        "HKLM:\SYSTEM\CurrentControlSet\Services\AdskLicensingService",
        "HKLM:\SYSTEM\CurrentControlSet\Services\FNPLicensingService",
        "HKLM:\SYSTEM\CurrentControlSet\Services\FlexNet Licensing Service",
        "HKLM:\SYSTEM\CurrentControlSet\Services\FlexNet Licensing Service 64"
    ) | ForEach-Object { Remove-RegistryKeySilently $_ }

    # Remove stale Autodesk entries from Programs list
    @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ) | ForEach-Object {
        Get-ChildItem -Path $_ -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
            if (($entry.DisplayName -like "*Autodesk*") -or ($entry.Publisher -like "*Autodesk*")) {
                Write-Host "  Removing stale entry: $($entry.DisplayName)" -ForegroundColor DarkGray
                Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────
# PHASE 7 — Temp folder general cleanup
# ─────────────────────────────────────────────────────────────
Write-Step "=== PHASE 7: General Temp Folder Cleanup ===" "Magenta"
Get-ChildItem -Path $Env:TEMP         -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path "C:\Windows\Temp" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "  Temp folders cleaned." -ForegroundColor DarkGray

# ─────────────────────────────────────────────────────────────
# DONE
# ─────────────────────────────────────────────────────────────
Clear-Host
Write-Host @"

╔══════════════════════════════════════════════════════════════╗
║       Autodesk Complete Removal — DONE                       ║
╠══════════════════════════════════════════════════════════════╣
║  Phase 1 : Products uninstalled (4 passes)                   ║
║  Phase 2 : Background services removed                       ║
║  Phase 3 : Leftover files and folders deleted                ║
║  Phase 4 : FLEXnet and licensing data wiped                  ║
║  Phase 5 : Telemetry, usage stats and logs removed           ║
║  Phase 6 : Registry cleaned (backup saved to Desktop)        ║
║  Phase 7 : Temp folders cleaned                              ║
╠══════════════════════════════════════════════════════════════╣
║  ► PLEASE RESTART YOUR COMPUTER before reinstalling.         ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green

if ($AutoRestart) {
    Write-Host "Restarting in 30 seconds... Press Ctrl+C to cancel." -ForegroundColor Red
    Start-Sleep -Seconds 30
    Restart-Computer -Force
}
