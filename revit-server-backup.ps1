#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Revit Server RVT Backup Script v2
    Runs on the Revit Server HOST machine.
    Uses revitservertool.exe createLocalRVT to export real .rvt files.

.FIXES in v2
    - Each .rvt file is placed inside its matching subfolder (mirrors Revit Server tree)
    - Locked models are detected and skipped gracefully (no crash, logged in manifest)

.NOTES
    - Run AS ADMINISTRATOR on the Revit Server host
    - revitservertool.exe is auto-detected from the Revit Server installation
    - Supported versions: 2020-2027
    - Output: Desktop\RevitServer_RVT_Backup\YYYY-MM-DD\VERSION_HOST_HH-MM\
#>

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
function Write-Title {
    param([string]$Text)
    $line = "=" * 64
    Write-Host ""
    Write-Host $line              -ForegroundColor Cyan
    Write-Host "  $Text"         -ForegroundColor Cyan
    Write-Host $line              -ForegroundColor Cyan
    Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]   $M" -ForegroundColor Green   }
function Write-Info { param([string]$M) Write-Host "  [..]  $M"  -ForegroundColor DarkGray }
function Write-Warn { param([string]$M) Write-Host "  [!!]  $M"  -ForegroundColor Yellow  }
function Write-Fail { param([string]$M) Write-Host "  [XX]  $M"  -ForegroundColor Red     }
function Write-Skip { param([string]$M) Write-Host "  [--]  $M"  -ForegroundColor DarkYellow }

# ─────────────────────────────────────────────────────────────
# STEP 1 — Detect installed Revit Server versions
# ─────────────────────────────────────────────────────────────
Write-Title "Revit Server RVT Backup v2 — Running on: $($env:COMPUTERNAME)"

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

Write-Host "  Scanning for installed Revit Server versions..." -ForegroundColor White
Write-Host ""

foreach ($v in $supportedVersions) {
    $toolExe  = $null
    $projRoot = $null
    $rsnIni   = $null

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
        "C:\ProgramData\Autodesk\Autodesk Revit Server $v\Config\RSN.ini"
    )) {
        if (Test-Path $r) { $rsnIni = $r; break }
    }

    if ($toolExe -or $projRoot) {
        $detectedVersions.Add([PSCustomObject]@{
            Version      = "$v"
            ToolExe      = $toolExe
            ProjectsRoot = $projRoot
            RSNIni       = $rsnIni
            HasTool      = ($null -ne $toolExe)
            HasProjects  = ($null -ne $projRoot)
        })
        $toolTag = if ($toolExe)  { "[TOOL OK]"     } else { "[NO TOOL]"     }
        $projTag = if ($projRoot) { "[PROJECTS OK]" } else { "[NO PROJECTS]" }
        Write-Host "  ► Revit Server $v  $toolTag  $projTag" -ForegroundColor Green
    }
}

if ($detectedVersions.Count -eq 0) {
    Write-Fail "No Revit Server installation detected on this machine."
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
    $choice = Read-Host "  Enter number"
    $idx    = [int]$choice - 1
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
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}',$version))" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualTool = Read-Host "  Enter full path manually (or press Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Tool not found. Exiting."
        exit 1
    }
    $toolExe = $manualTool.Trim()
}

$toolVersion = (Get-Item $toolExe).VersionInfo.FileVersion
Write-OK "Tool     : $toolExe"
Write-OK "Version  : $toolVersion"

# ─────────────────────────────────────────────────────────────
# STEP 4 — Determine server hostname
# ─────────────────────────────────────────────────────────────
Write-Title "Step 4: Determining Server Host"

$serverHost = $env:COMPUTERNAME

if ($rsnIniPath -and (Test-Path $rsnIniPath)) {
    $rsnContent = Get-Content $rsnIniPath -ErrorAction SilentlyContinue
    $rsnHost    = ($rsnContent | Where-Object { $_ -match '\S' } | Select-Object -First 1).Trim()
    if ($rsnHost) {
        $serverHost = $rsnHost
        Write-OK "Host from RSN.ini : $serverHost"
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

# Build model list
# Revit Server tree example:
#   Projects\Viktoriia Ilytska\Viktoriia Ilytska.rvt   (folder on disk)
# RSN path:
#   Viktoriia Ilytska/Viktoriia Ilytska.rvt
# Target backup structure:
#   backup\Viktoriia Ilytska\Viktoriia Ilytska.rvt     (real .rvt file)

$models = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($f in $modelFolders) {
    $relPath   = $f.FullName.Replace($projRoot, "").TrimStart("\")
    $rsnPath   = $relPath.Replace("\", "/")
    # The DESTINATION folder = the parent folder of the model, mirroring the tree
    # e.g. "Viktoriia Ilytska\Viktoriia Ilytska.rvt" -> destSubFolder = "Viktoriia Ilytska"
    $destSubFolder = [System.IO.Path]::GetDirectoryName($relPath)

    $models.Add([PSCustomObject]@{
        Name          = $f.Name              # Viktoriia Ilytska.rvt
        RelPath       = $relPath             # Viktoriia Ilytska\Viktoriia Ilytska.rvt
        RSNPath       = $rsnPath             # Viktoriia Ilytska/Viktoriia Ilytska.rvt
        DestSubFolder = $destSubFolder       # Viktoriia Ilytska
    })
}

Write-Host "  Model tree:" -ForegroundColor White
foreach ($m in $models) {
    Write-Host "    RSN://$serverHost/$($m.RSNPath)" -ForegroundColor DarkCyan
}

# ─────────────────────────────────────────────────────────────
# STEP 7 — Create backup destination
# ─────────────────────────────────────────────────────────────
Write-Title "Step 7: Creating Backup Folder"

$desktopPath = [Environment]::GetFolderPath("Desktop")
$dateStamp   = Get-Date -Format "yyyy-MM-dd"
$timeStamp   = Get-Date -Format "HH-mm"
$backupDest  = Join-Path $desktopPath "RevitServer_RVT_Backup\$dateStamp\${version}_${serverHost}_${timeStamp}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null
Write-OK "Backup root: $backupDest"

# ─────────────────────────────────────────────────────────────
# STEP 8 — Export RVT files
# ─────────────────────────────────────────────────────────────
Write-Title "Step 8: Exporting RVT Files"

Write-Host "  Locked models will be SKIPPED and logged." -ForegroundColor Yellow
Write-Host ""

$successList = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$failList    = [System.Collections.Generic.List[string]]::new()
$current     = 0

foreach ($m in $models) {
    $current++

    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    # ── Create the EXACT subfolder matching the Revit Server tree ──────────
    # If model is at root level (no subfolder) -> goes directly in backupDest
    # If model is at  SubFolder\Model.rvt      -> goes in backupDest\SubFolder\
    $targetDir = if ([string]::IsNullOrWhiteSpace($m.DestSubFolder)) {
        $backupDest
    } else {
        Join-Path $backupDest $m.DestSubFolder
    }

    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    Write-Info "  Destination: $targetDir\$($m.Name)"

    # revitservertool.exe createLocalRVT "RSNPath" -s Host -d "DestFolder" -o
    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$targetDir`" -o"

    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode

        switch ($exitCode) {
            0 {
                Write-OK "Done: $($m.Name)"
                $successList.Add($m.RSNPath)
            }
            # Exit code 5 = access denied (model locked / in use by a user)
            5 {
                Write-Skip "SKIPPED (model locked by a user): $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [locked / access denied]")
            }
            # Exit code 1 = model busy / exclusive access by Revit Server
            1 {
                Write-Skip "SKIPPED (model busy): $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [model busy]")
            }
            default {
                Write-Fail "FAILED (exit $exitCode): $($m.Name)"
                $failList.Add("$($m.RSNPath)  [exit code $exitCode]")
            }
        }
    } catch {
        Write-Fail "Exception: $($m.Name) — $($_.Exception.Message)"
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
$manifest.Add("Revit Server RVT Backup Manifest v2")
$manifest.Add("=" * 64)
$manifest.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$manifest.Add("Revit Version : $version")
$manifest.Add("Server Host   : $serverHost")
$manifest.Add("Machine       : $($env:COMPUTERNAME)")
$manifest.Add("Tool          : $toolExe")
$manifest.Add("Projects Root : $projRoot")
$manifest.Add("Destination   : $backupDest")
$manifest.Add("Models total  : $modelCount")
$manifest.Add("Succeeded     : $successCount")
$manifest.Add("Skipped       : $skipCount  (locked/busy - not an error)")
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
    $manifest.Add("  Tip: run backup again after business hours when no users are in the models.")
    $manifest.Add("")
}

if ($failList.Count -gt 0) {
    $manifest.Add("FAILED ($failCount):")
    $manifest.Add("-" * 64)
    foreach ($f in $failList) { $manifest.Add("  [XX] $f") }
    $manifest.Add("")
}

$manifest | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest: $manifestPath"

# ─────────────────────────────────────────────────────────────
# DONE — Summary
# ─────────────────────────────────────────────────────────────
Write-Title "Backup Complete"

Write-Host "  Server        : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Total models  : $modelCount"   -ForegroundColor White
Write-Host "  Succeeded     : $successCount" -ForegroundColor Green
Write-Host "  Skipped       : $skipCount"    -ForegroundColor $(if ($skipCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Failed        : $failCount"    -ForegroundColor $(if ($failCount -gt 0) { "Red"    } else { "Green" })
Write-Host ""
Write-Host "  Location  : $backupDest" -ForegroundColor Cyan
Write-Host "  Manifest  : $manifestPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Structure:" -ForegroundColor DarkGray
Write-Host "  RevitServer_RVT_Backup\"                                   -ForegroundColor DarkCyan
Write-Host "  └── $dateStamp\"                                            -ForegroundColor DarkCyan
Write-Host "      └── ${version}_${serverHost}_${timeStamp}\"            -ForegroundColor DarkCyan
Write-Host "          ├── _BACKUP_MANIFEST.txt"                          -ForegroundColor DarkGray
foreach ($m in ($models | Select-Object -First 6)) {
    if ($m.DestSubFolder) {
        Write-Host "          ├── $($m.DestSubFolder)\"                   -ForegroundColor DarkCyan
        Write-Host "          │   └── $($m.Name)"                        -ForegroundColor DarkGray
    } else {
        Write-Host "          ├── $($m.Name)"                            -ForegroundColor DarkGray
    }
}
Write-Host ""

$open = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($open.Trim().ToUpper() -eq "Y") { Start-Process explorer.exe $backupDest }

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
