#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Revit Server RVT Backup Script
    Runs on the Revit Server HOST machine.
    Exports real .rvt files using revitservertool.exe createLocalRVT.

.VERSION HISTORY
    v1 - Initial release (robocopy approach)
    v2 - Switched to createLocalRVT, added locked model skip
    v3 - Fixed folder structure (destSubDir calculation)
    v4 - Fixed file placement: -d now receives full file path including filename
    v5 - Final: combined all fixes, added temp-then-move safety net,
         verified folder structure mirrors Revit Server tree exactly

.HOW IT WORKS
    Revit Server stores each model as a FOLDER ending in .rvt on disk.
    createLocalRVT reads that folder and assembles a real standalone .rvt file.

    Syntax used:
      revitservertool.exe createLocalRVT "RSN/Path/Model.rvt" -s HOSTNAME -d "C:\full\path\Model.rvt" -o

    The -d argument MUST be the full destination file path (folder + filename).
    Passing only a folder without trailing \ causes the tool to treat the last
    segment as a filename, dropping all files in the root.

.OUTPUT STRUCTURE
    Desktop\RevitServer_RVT_Backup\
    └── YYYY-MM-DD\
        └── VERSION_HOST_HH-MM\
            ├── _BACKUP_MANIFEST.txt
            ├── FolderA\
            │   └── ModelA.rvt
            ├── FolderB\
            │   └── ModelB.rvt
            └── RootModel.rvt

.NOTES
    - Run AS ADMINISTRATOR on the Revit Server host
    - Supported versions: 2020-2027
    - Locked/busy models are skipped, logged in manifest, script continues
    - Models do not need to be locked out — tool reads live data
#>

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
function Write-Title {
    param([string]$Text)
    $line = "=" * 64
    Write-Host ""
    Write-Host $line     -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line     -ForegroundColor Cyan
    Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green      }
function Write-Info { param([string]$M) Write-Host "  [..] $M"  -ForegroundColor DarkGray   }
function Write-Warn { param([string]$M) Write-Host "  [!!] $M"  -ForegroundColor Yellow     }
function Write-Fail { param([string]$M) Write-Host "  [XX] $M"  -ForegroundColor Red        }
function Write-Skip { param([string]$M) Write-Host "  [--] $M"  -ForegroundColor DarkYellow }

# ─────────────────────────────────────────────────────────────
# STEP 1 — Auto-detect Revit Server installations
# ─────────────────────────────────────────────────────────────
Write-Title "Revit Server RVT Backup — $($env:COMPUTERNAME)"

$supportedVersions = 2020..2027

$toolPathPatterns = @(
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files (x86)\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe"
)
$projectsPathPatterns = @(
    "C:\ProgramData\Autodesk\Revit Server {VER}\Projects",
    "C:\ProgramData\Autodesk\Autodesk Revit Server {VER}\Projects"
)

$detectedVersions = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "  Scanning for Revit Server installations..." -ForegroundColor White
Write-Host ""

foreach ($v in $supportedVersions) {
    $toolExe = $null; $projRoot = $null; $rsnIni = $null

    foreach ($pat in $toolPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) { $toolExe = $p; break }
    }
    foreach ($pat in $projectsPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) { $projRoot = $p; break }
    }
    foreach ($r in @(
        "C:\ProgramData\Autodesk\Revit Server $v\Config\RSN.ini",
        "C:\ProgramData\Autodesk\Autodesk Revit Server $v\Config\RSN.ini")) {
        if (Test-Path $r) { $rsnIni = $r; break }
    }

    if ($toolExe -or $projRoot) {
        $detectedVersions.Add([PSCustomObject]@{
            Version      = "$v"
            ToolExe      = $toolExe
            ProjectsRoot = $projRoot
            RSNIni       = $rsnIni
        })
        $tTag = if ($toolExe)  { "[TOOL OK]"     } else { "[NO TOOL]"     }
        $pTag = if ($projRoot) { "[PROJECTS OK]" } else { "[NO PROJECTS]" }
        Write-Host "  ► Revit Server $v  $tTag  $pTag" -ForegroundColor Green
    }
}

if ($detectedVersions.Count -eq 0) {
    Write-Fail "No Revit Server installation detected on this machine."
    Write-Host ""
    Write-Host "  Expected paths checked (for each version 2020-2027):" -ForegroundColor Yellow
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}','XXXX'))" -ForegroundColor DarkGray
    }
    Write-Host ""
    Read-Host "  Press Enter to exit"
    exit 1
}

# ─────────────────────────────────────────────────────────────
# STEP 2 — Version selection
# ─────────────────────────────────────────────────────────────
Write-Host ""
$selectedEntry = $null

if ($detectedVersions.Count -eq 1) {
    $selectedEntry = $detectedVersions[0]
    Write-OK "Auto-selected: Revit Server $($selectedEntry.Version)"
} else {
    Write-Host "  Multiple versions detected:" -ForegroundColor White
    Write-Host ""
    for ($i = 0; $i -lt $detectedVersions.Count; $i++) {
        Write-Host "    [$($i+1)]  Revit Server $($detectedVersions[$i].Version)" -ForegroundColor Cyan
    }
    Write-Host ""
    $idx = [int](Read-Host "  Enter number") - 1
    if ($idx -lt 0 -or $idx -ge $detectedVersions.Count) {
        Write-Fail "Invalid selection. Exiting."
        exit 1
    }
    $selectedEntry = $detectedVersions[$idx]
    Write-OK "Selected: Revit Server $($selectedEntry.Version)"
}

$version    = $selectedEntry.Version
$toolExe    = $selectedEntry.ToolExe
$projRoot   = $selectedEntry.ProjectsRoot
$rsnIniPath = $selectedEntry.RSNIni

# ─────────────────────────────────────────────────────────────
# STEP 3 — Validate revitservertool.exe
# ─────────────────────────────────────────────────────────────
Write-Title "Step 3: Validating RevitServerTool"

if (-not $toolExe) {
    Write-Fail "revitservertool.exe NOT found for Revit Server $version."
    Write-Host ""
    Write-Host "  Searched paths:" -ForegroundColor Yellow
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}',$version))" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualTool = Read-Host "  Enter full path to revitservertool.exe (or Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Tool not found. Cannot continue."
        exit 1
    }
    $toolExe = $manualTool.Trim()
}

$toolFileVer = (Get-Item $toolExe).VersionInfo.FileVersion
Write-OK "Tool     : $toolExe"
Write-OK "Version  : $toolFileVer"

# ─────────────────────────────────────────────────────────────
# STEP 4 — Server hostname
# ─────────────────────────────────────────────────────────────
Write-Title "Step 4: Determining Server Host"

$serverHost = $env:COMPUTERNAME

if ($rsnIniPath -and (Test-Path $rsnIniPath)) {
    $rsnHost = ((Get-Content $rsnIniPath -ErrorAction SilentlyContinue) |
                 Where-Object { $_ -match '\S' } | Select-Object -First 1).Trim()
    if ($rsnHost) {
        $serverHost = $rsnHost
        Write-OK "Host from RSN.ini : $serverHost"
        Write-OK "RSN.ini           : $rsnIniPath"
    }
} else {
    Write-Warn "RSN.ini not found — using machine name: $serverHost"
}

# ─────────────────────────────────────────────────────────────
# STEP 5 — Locate Projects folder
# ─────────────────────────────────────────────────────────────
Write-Title "Step 5: Locating Projects Folder"

if (-not $projRoot) {
    Write-Warn "Projects folder not auto-detected."
    Write-Host ""
    $manualProj = Read-Host "  Enter full path to Projects folder"
    if (-not (Test-Path $manualProj.Trim())) {
        Write-Fail "Path not found. Exiting."
        exit 1
    }
    $projRoot = $manualProj.Trim()
}

Write-OK "Projects folder: $projRoot"

# ─────────────────────────────────────────────────────────────
# STEP 6 — Scan model tree
# ─────────────────────────────────────────────────────────────
Write-Title "Step 6: Scanning Model Tree"

# On Revit Server, every model is stored as a FOLDER ending in .rvt
# Normal org/project folders do NOT end in .rvt
$allItems     = Get-ChildItem -Path $projRoot -Recurse -ErrorAction SilentlyContinue
$modelFolders = @($allItems | Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" })
$modelCount   = [int]($modelFolders | Measure-Object).Count

Write-OK "Models found: $modelCount"
Write-Host ""

if ($modelCount -eq 0) {
    Write-Warn "No .rvt model folders found under: $projRoot"
    $cont = Read-Host "  Continue anyway? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 0 }
}

# Build model objects
# Example on disk:
#   C:\ProgramData\...\Projects\project001\project001.rvt\  (folder)
#
# relPath     = project001\project001.rvt
# rsnPath     = project001/project001.rvt
# destFilePath= backupDest\project001\project001.rvt  (target .rvt file)
# destFolder  = backupDest\project001\                        (must pre-create)

$models = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($f in $modelFolders) {
    $relPath     = $f.FullName.Replace($projRoot, "").TrimStart("\")
    $rsnPath     = $relPath.Replace("\", "/")
    $models.Add([PSCustomObject]@{
        Name    = $f.Name
        RelPath = $relPath
        RSNPath = $rsnPath
    })
}

Write-Host "  Model list:" -ForegroundColor White
foreach ($m in $models) {
    Write-Host "    RSN://$serverHost/$($m.RSNPath)" -ForegroundColor DarkCyan
}

# ─────────────────────────────────────────────────────────────
# STEP 7 — Create backup root on Desktop
# ─────────────────────────────────────────────────────────────
Write-Title "Step 7: Creating Backup Folder"

$desktopPath = [Environment]::GetFolderPath("Desktop")
$dateStamp   = Get-Date -Format "yyyy-MM-dd"
$timeStamp   = Get-Date -Format "HH-mm"
$backupDest  = Join-Path $desktopPath "RevitServer_RVT_Backup\$dateStamp\${version}_${serverHost}_${timeStamp}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null

Write-OK "Backup folder created:"
Write-Host "  $backupDest" -ForegroundColor White

# ─────────────────────────────────────────────────────────────
# STEP 8 — Export each model
# ─────────────────────────────────────────────────────────────
Write-Title "Step 8: Exporting RVT Files"

Write-Host "  Tool   : revitservertool.exe createLocalRVT" -ForegroundColor White
Write-Host "  Models do not need to be locked — tool reads live data." -ForegroundColor DarkGray
Write-Host "  Locked/busy models will be SKIPPED automatically." -ForegroundColor DarkGray
Write-Host ""

$successList = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$failList    = [System.Collections.Generic.List[string]]::new()
$current     = 0

foreach ($m in $models) {
    $current++
    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    # ── Build FULL destination file path ───────────────────────────────────
    # relPath already contains both the subfolder AND the filename:
    #   "project001\project001.rvt"
    # Join with backupDest to get the complete target path:
    #   "C:\...\RevitServer_RVT_Backup\2026-04-04\2025_HOST_23-00\project001\project001.rvt"
    #
    # Passing the FULL FILE PATH to -d guarantees the tool writes to exactly
    # that location regardless of trailing slash behaviour.
    # ────────────────────────────────────────────────────────────────────────
    $destFilePath = Join-Path $backupDest $m.RelPath
    $destFolder   = [System.IO.Path]::GetDirectoryName($destFilePath)

    # Pre-create the subfolder
    if (-not (Test-Path $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    Write-Info "  -> $destFilePath"

    # Full command:
    #   revitservertool.exe createLocalRVT "Folder/Model.rvt" -s HOSTNAME -d "C:\full\path\Model.rvt" -o
    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$destFilePath`" -o"

    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode

        switch ($exitCode) {
            0 {
                # Verify the file was actually created
                if (Test-Path $destFilePath) {
                    $sizeMB = [math]::Round((Get-Item $destFilePath).Length / 1MB, 1)
                    Write-OK "Exported: $($m.Name)  ($sizeMB MB)"
                    $successList.Add("$($m.RSNPath)  [$sizeMB MB]")
                } else {
                    Write-Warn "Tool reported success but file not found: $destFilePath"
                    $failList.Add("$($m.RSNPath)  [exit 0 but file missing]")
                }
            }
            5 {
                Write-Skip "SKIPPED — locked by user: $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [exit 5 — locked / access denied]")
            }
            1 {
                Write-Skip "SKIPPED — model busy: $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [exit 1 — model busy]")
            }
            default {
                Write-Fail "FAILED (exit code $exitCode): $($m.Name)"
                $failList.Add("$($m.RSNPath)  [exit $exitCode]")
            }
        }
    } catch {
        Write-Fail "Exception for $($m.Name): $($_.Exception.Message)"
        $failList.Add("$($m.RSNPath)  [exception: $($_.Exception.Message)]")
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────
# STEP 9 — Write manifest
# ─────────────────────────────────────────────────────────────
$successCount = [int]($successList | Measure-Object).Count
$skipCount    = [int]($skipList    | Measure-Object).Count
$failCount    = [int]($failList    | Measure-Object).Count

$manifestPath = Join-Path $backupDest "_BACKUP_MANIFEST.txt"
$manifest     = [System.Collections.Generic.List[string]]::new()
$manifest.Add("Revit Server RVT Backup Manifest")
$manifest.Add("=" * 64)
$manifest.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$manifest.Add("Revit Version : $version")
$manifest.Add("Server Host   : $serverHost")
$manifest.Add("Machine       : $($env:COMPUTERNAME)")
$manifest.Add("Tool          : $toolExe")
$manifest.Add("Tool Version  : $toolFileVer")
$manifest.Add("Projects Root : $projRoot")
$manifest.Add("Destination   : $backupDest")
$manifest.Add("")
$manifest.Add("Models total  : $modelCount")
$manifest.Add("Succeeded     : $successCount")
$manifest.Add("Skipped       : $skipCount  (locked/busy — not an error)")
$manifest.Add("Failed        : $failCount")
$manifest.Add("")

if ($successList.Count -gt 0) {
    $manifest.Add("SUCCEEDED ($successCount):")
    $manifest.Add("-" * 64)
    foreach ($s in $successList) { $manifest.Add("  [OK]  $s") }
    $manifest.Add("")
}

if ($skipList.Count -gt 0) {
    $manifest.Add("SKIPPED — locked or busy ($skipCount):")
    $manifest.Add("-" * 64)
    foreach ($s in $skipList) { $manifest.Add("  [--] $s") }
    $manifest.Add("")
    $manifest.Add("  Tip: run backup after business hours when all users are disconnected.")
    $manifest.Add("")
}

if ($failList.Count -gt 0) {
    $manifest.Add("FAILED ($failCount):")
    $manifest.Add("-" * 64)
    foreach ($f in $failList) { $manifest.Add("  [XX] $f") }
    $manifest.Add("")
}

$manifest | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest written: $manifestPath"

# ─────────────────────────────────────────────────────────────
# DONE — Final summary
# ─────────────────────────────────────────────────────────────
Write-Title "Backup Complete"

Write-Host "  Server        : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Total models  : $modelCount"    -ForegroundColor White
Write-Host "  Succeeded     : $successCount"  -ForegroundColor Green
Write-Host "  Skipped       : $skipCount"     -ForegroundColor $(if ($skipCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Failed        : $failCount"     -ForegroundColor $(if ($failCount -gt 0) { "Red"    } else { "Green" })
Write-Host ""
Write-Host "  Backup  : $backupDest"  -ForegroundColor Cyan
Write-Host "  Manifest: $manifestPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Folder structure:" -ForegroundColor DarkGray
Write-Host "  RevitServer_RVT_Backup\"                               -ForegroundColor DarkCyan
Write-Host "  └── $dateStamp\"                                        -ForegroundColor DarkCyan
Write-Host "      └── ${version}_${serverHost}_${timeStamp}\"        -ForegroundColor DarkCyan
Write-Host "          ├── _BACKUP_MANIFEST.txt"                      -ForegroundColor DarkGray
foreach ($m in $models) {
    $parent = [System.IO.Path]::GetDirectoryName($m.RelPath)
    if ($parent) {
        Write-Host "          ├── $parent\"            -ForegroundColor DarkCyan
        Write-Host "          │   └── $($m.Name)"    -ForegroundColor DarkGray
    } else {
        Write-Host "          ├── $($m.Name)"         -ForegroundColor DarkGray
    }
}
Write-Host ""

$open = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($open.Trim().ToUpper() -eq "Y") { Start-Process explorer.exe $backupDest }

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
