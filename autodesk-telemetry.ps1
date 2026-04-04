#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs a Windows Scheduled Task that runs daily to wipe
    all Autodesk telemetry, usage statistics, analytics, logs,
    and cached data from the local machine.

.NOTES
    Run this ONCE as Administrator to register the task.
    The cleanup will then run automatically every day at 03:00.
    Task name: "Autodesk Daily Telemetry Cleanup"
#>

# ─────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────
$TaskName        = "Autodesk Daily Telemetry Cleanup"
$TaskDescription = "Clears Autodesk telemetry, analytics, usage stats, logs and cached data daily."
$RunTime         = "03:00"   # 24h format — change if needed
# ─────────────────────────────────────────────────────────────

$CleanupScript = @'
function Remove-PathSilently {
    param([string]$Path)
    if (Test-Path $Path) {
        try { Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop }
        catch {}
    }
}

# ── Per-machine paths ────────────────────────────────────────
@(
    "C:\ProgramData\Autodesk\ADLM",
    "C:\ProgramData\Autodesk\ODIS\cache",
    "C:\ProgramData\Autodesk\ODIS\logs",
    "C:\ProgramData\FLEXnet",
    "C:\Windows\Temp\Autodesk",
    "C:\ProgramData\Autodesk\CER",
    "C:\ProgramData\Autodesk\Desktop App",
    "C:\ProgramData\Autodesk\Analytics"
) | ForEach-Object { Remove-PathSilently $_ }

# ── Per-user paths (all profiles on machine) ─────────────────
Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $base    = $_.FullName
    $appdata = "$base\AppData\Roaming"
    $local   = "$base\AppData\Local"

    # ── Years loop 2020–2027 ──────────────────────────────────
    $years = 2020..2027

    $paths = @(
        # ADLM, Analytics, CER, Licensing
        "$appdata\Autodesk\ADLM",
        "$local\Autodesk\ADLM",
        "$local\Autodesk\Analytics",
        "$local\Autodesk\Web Services",
        "$appdata\Autodesk\CER",
        "$local\Autodesk\CER",
        "$local\Autodesk Desktop App\logs",
        "$local\Autodesk Desktop App\cache",
        "$appdata\Autodesk Desktop App\logs",
        "$local\Autodesk\Genuine Autodesk Service",
        "$local\Autodesk\AdskLicensing\logs",
        "$local\Autodesk\AdskLicensing\cache",
        "$local\Temp\Autodesk"
    )

    foreach ($year in $years) {
        $paths += @(
            # Revit journals & logs
            "$appdata\Autodesk\Revit\Autodesk Revit $year\Journals",
            "$appdata\Autodesk\Revit\Autodesk Revit $year\logs",
            "$local\Autodesk\Revit\Autodesk Revit $year\CollaborationCache",

            # AutoCAD cache & logs
            "$local\Autodesk\AutoCAD $year\en-US\Cache",
            "$local\Autodesk\AutoCAD $year\en-US\logs",
            "$appdata\Autodesk\AutoCAD $year\en-US\Cache",

            # AutoCAD LT
            "$local\Autodesk\AutoCAD LT $year\en-US\Cache",
            "$appdata\Autodesk\AutoCAD LT $year\en-US\Cache",

            # Civil 3D
            "$local\Autodesk\Civil 3D $year",
            "$appdata\Autodesk\Civil 3D $year",

            # Revit Server cache
            "$local\Autodesk\Revit Server $year",

            # Navisworks cache & logs
            "$appdata\Autodesk\Navisworks Manage $year",
            "$appdata\Autodesk\Navisworks Simulate $year",
            "$appdata\Autodesk\Navisworks Freedom $year",
            "$local\Autodesk\Navisworks Manage $year",
            "$local\Autodesk\Navisworks Simulate $year",

            # Inventor cache & logs
            "$appdata\Autodesk\Inventor $year",
            "$local\Autodesk\Inventor $year",

            # 3ds Max logs & cache
            "$local\Autodesk\3dsMax $year",
            "$appdata\Autodesk\3dsMax $year",
            "$local\Autodesk\3ds Max $year",
            "$appdata\Autodesk\3ds Max $year",

            # Maya logs & cache
            "$appdata\Autodesk\Maya\$year",
            "$local\Autodesk\Maya\$year",
            "$appdata\Autodesk\Maya$year",

            # InfraWorks
            "$local\Autodesk\InfraWorks $year",
            "$appdata\Autodesk\InfraWorks $year",

            # ReCap
            "$local\Autodesk\ReCap $year",
            "$appdata\Autodesk\ReCap $year",

            # Vault
            "$appdata\Autodesk\Vault $year",
            "$local\Autodesk\Vault $year",

            # Fusion 360 (year-versioned logs)
            "$appdata\Autodesk\Autodesk Fusion 360\logs\$year",

            # Structural / Advance Steel
            "$appdata\Autodesk\Advance Steel $year",
            "$local\Autodesk\Advance Steel $year",

            # Map 3D
            "$appdata\Autodesk\AutoCAD Map 3D $year",
            "$local\Autodesk\AutoCAD Map 3D $year",

            # Plant 3D
            "$appdata\Autodesk\AutoCAD Plant 3D $year",
            "$local\Autodesk\AutoCAD Plant 3D $year",

            # MEP
            "$appdata\Autodesk\AutoCAD MEP $year",
            "$local\Autodesk\AutoCAD MEP $year",

            # Architecture
            "$appdata\Autodesk\AutoCAD Architecture $year",
            "$local\Autodesk\AutoCAD Architecture $year",

            # Electrical
            "$appdata\Autodesk\AutoCAD Electrical $year",
            "$local\Autodesk\AutoCAD Electrical $year",

            # Mechanical
            "$appdata\Autodesk\AutoCAD Mechanical $year",
            "$local\Autodesk\AutoCAD Mechanical $year",

            # BIM Collaborate / BIM 360
            "$local\Autodesk\BIM 360\$year",
            "$appdata\Autodesk\BIM 360\$year",

            # Dynamo
            "$appdata\Dynamo\$year",
            "$local\Dynamo\$year",

            # Forma
            "$local\Autodesk\Forma\$year",
            "$appdata\Autodesk\Forma\$year"
        )
    }

    # Fusion 360 main logs (not year-versioned)
    $paths += @(
        "$appdata\Autodesk\Autodesk Fusion 360\logs",
        "$local\Autodesk\Autodesk Fusion 360\logs",
        "$local\Autodesk\Autodesk Fusion 360\cache"
    )

    foreach ($p in $paths) { Remove-PathSilently $p }
}

# ── Wildcard log / temp file cleanup ────────────────────────
@($Env:TEMP, "C:\Windows\Temp") | ForEach-Object {
    if (Test-Path $_) {
        Get-ChildItem $_ -Filter "*.adsklog"     -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem $_ -Filter "Adsk*"         -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Get-ChildItem $_ -Filter "autodesk*"     -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Get-ChildItem $_ -Filter "AdskIdentity*" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ── Write completion log (last 30 entries kept) ──────────────
$logDir = "C:\ProgramData\Autodesk\CleanupLog"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logFile = "$logDir\cleanup.log"
Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Autodesk telemetry cleanup completed." -Encoding UTF8
$lines = Get-Content $logFile -ErrorAction SilentlyContinue
if ($lines.Count -gt 30) { $lines | Select-Object -Last 30 | Set-Content $logFile -Encoding UTF8 }
'@

# ── Save the cleanup script to ProgramData ───────────────────
$scriptDir  = "C:\ProgramData\AutodeskCleanup"
$scriptPath = "$scriptDir\autodesk-telemetry-cleanup.ps1"

if (-not (Test-Path $scriptDir)) { New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null }
$CleanupScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
Write-Host "Cleanup script saved to: $scriptPath" -ForegroundColor Cyan

# ── Build and register the Scheduled Task ────────────────────
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

$trigger = New-ScheduledTaskTrigger -Daily -At $RunTime

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false `
    -MultipleInstances IgnoreNew

$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description $TaskDescription `
    -Action      $action `
    -Trigger     $trigger `
    -Settings    $settings `
    -Principal   $principal `
    -Force | Out-Null

Write-Host ""
Write-Host "Task registered! Runs daily at $RunTime as SYSTEM."               -ForegroundColor Green
Write-Host "Cleanup script : $scriptPath"                                      -ForegroundColor Cyan
Write-Host "Log file       : C:\ProgramData\Autodesk\CleanupLog\cleanup.log"  -ForegroundColor Cyan
Write-Host ""
Write-Host "Run it NOW  : Start-ScheduledTask -TaskName '$TaskName'"           -ForegroundColor Yellow
Write-Host "Remove task : Unregister-ScheduledTask -TaskName '$TaskName'"      -ForegroundColor DarkYellow
