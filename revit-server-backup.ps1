#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Revit Server RVT Backup Script
    Runs on the Revit Server HOST machine.
    Uses revitservertool.exe createLocalRVT to export real .rvt files.

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
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green   }
function Write-Info { param([string]$M) Write-Host "  [..] $M"  -ForegroundColor DarkGray }
function Write-Warn { param([string]$M) Write-Host "  [!!] $M"  -ForegroundColor Yellow  }
function Write-Fail { param([string]$M) Write-Host "  [XX] $M"  -ForegroundColor Red     }

# ─────────────────────────────────────────────────────────────
# STEP 1 — Detect installed Revit Server versions
# ─────────────────────────────────────────────────────────────
Write-Title "Revit Server RVT Backup — Running on: $($env:COMPUTERNAME)"

$supportedVersions = 2020..2027

# Tool path patterns — Autodesk changed the folder name between versions
# Each pattern is tried in order for every version
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
    $toolExe   = $null
    $projRoot  = $null
    $rsnIni    = $null

    foreach ($pat in $toolPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) { $toolExe = $p; break }
    }
    foreach ($pat in $projectsPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) { $projRoot = $p; break }
    }

    $rsnCandidates = @(
        "C:\ProgramData\Autodesk\Revit Server $v\Config\RSN.ini",
        "C:\ProgramData\Autodesk\Autodesk Revit Server $v\Config\RSN.ini"
    )
    foreach ($r in $rsnCandidates) {
        if (Test-Path $r) { $rsnIni = $r; break }
    }

    if ($toolExe -or $projRoot) {
        $detectedVersions.Add([PSCustomObject]@{
            Version    = "$v"
            ToolExe    = $toolExe
            ProjectsRoot = $projRoot
            RSNIni     = $rsnIni
            HasTool    = ($null -ne $toolExe)
            HasProjects= ($null -ne $projRoot)
        })
        $toolTag = if ($toolExe)  { "[TOOL OK]"    } else { "[NO TOOL]" }
        $projTag = if ($projRoot) { "[PROJECTS OK]"} else { "[NO PROJECTS]" }
        Write-Host "  ► Revit Server $v  $toolTag  $projTag" -ForegroundColor Green
    }
}

if ($detectedVersions.Count -eq 0) {
    Write-Fail "No Revit Server installation detected on this machine."
    Write-Host ""
    Write-Host "  Expected locations checked:" -ForegroundColor Yellow
    foreach ($pat in $toolPathPatterns) { Write-Host "    $($pat.Replace('{VER}','XXXX'))" -ForegroundColor DarkGray }
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
    Write-OK "Auto-selected only detected version: Revit Server $($selectedEntry.Version)"
} else {
    Write-Host "  Multiple versions detected. Which one do you want to back up?" -ForegroundColor White
    Write-Host ""
    for ($i = 0; $i -lt $detectedVersions.Count; $i++) {
        $e = $detectedVersions[$i]
        Write-Host "    [$($i+1)]  Revit Server $($e.Version)" -ForegroundColor Cyan
    }
    Write-Host ""
    $choice = Read-Host "  Enter number"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge $detectedVersions.Count) {
        Write-Fail "Invalid selection. Exiting."
        exit 1
    }
    $selectedEntry = $detectedVersions[$idx]
    Write-OK "Selected: Revit Server $($selectedEntry.Version)"
}

$version     = $selectedEntry.Version
$toolExe     = $selectedEntry.ToolExe
$projRoot    = $selectedEntry.ProjectsRoot
$rsnIniPath  = $selectedEntry.RSNIni

# ─────────────────────────────────────────────────────────────
# STEP 3 — Validate revitservertool.exe
# ─────────────────────────────────────────────────────────────
Write-Title "Step 3: Validating RevitServerTool"

if (-not $toolExe) {
    Write-Fail "revitservertool.exe NOT found for Revit Server $version."
    Write-Host ""
    Write-Host "  Expected paths:" -ForegroundColor Yellow
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}',$version))" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  The tool ships with Revit Server — check that the installation is complete." -ForegroundColor Yellow
    Write-Host "  You can also enter the path manually:" -ForegroundColor Yellow
    $manualTool = Read-Host "  Full path to revitservertool.exe (or press Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Tool not found. Cannot continue without revitservertool.exe."
        exit 1
    }
    $toolExe = $manualTool.Trim()
}

Write-OK "Tool      : $toolExe"

# Quick version test
$toolVersion = (Get-Item $toolExe).VersionInfo.FileVersion
Write-OK "File ver  : $toolVersion"

# ─────────────────────────────────────────────────────────────
# STEP 4 — Determine server hostname
# ─────────────────────────────────────────────────────────────
Write-Title "Step 4: Determining Server Host"

$serverHost = $env:COMPUTERNAME   # default — we ARE the server

if ($rsnIniPath -and (Test-Path $rsnIniPath)) {
    $rsnContent = Get-Content $rsnIniPath -ErrorAction SilentlyContinue
    $rsnHost    = ($rsnContent | Where-Object { $_ -match '\S' } | Select-Object -First 1).Trim()
    if ($rsnHost) {
        $serverHost = $rsnHost
        Write-OK "Host from RSN.ini : $serverHost"
        Write-OK "RSN.ini           : $rsnIniPath"
    }
} else {
    Write-Warn "RSN.ini not found — using local machine name: $serverHost"
}

# ─────────────────────────────────────────────────────────────
# STEP 5 — Locate and validate Projects folder
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

# On Revit Server, each model IS a folder ending in .rvt
# org folders are normal folders without .rvt extension
$allItems     = Get-ChildItem -Path $projRoot -Recurse -ErrorAction SilentlyContinue
$modelFolders = @($allItems | Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" })
$modelCount   = [int]($modelFolders | Measure-Object).Count

Write-OK "Model folders found: $modelCount"
Write-Host ""

if ($modelCount -eq 0) {
    Write-Warn "No .rvt model folders found under: $projRoot"
    $cont = Read-Host "  Continue anyway? (Y/N)"
    if ($cont.Trim().ToUpper() -ne "Y") { exit 0 }
}

# Build model list with RSN paths
$models = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($f in $modelFolders) {
    $relPath   = $f.FullName.Replace($projRoot, "").TrimStart("\")
    $rsnPath   = $relPath.Replace("\", "/")
    $subFolder = [System.IO.Path]::GetDirectoryName($relPath)
    $models.Add([PSCustomObject]@{
        Name      = $f.Name
        RelPath   = $relPath
        RSNPath   = $rsnPath
        SubFolder = $subFolder
    })
}

Write-Host "  Model list:" -ForegroundColor White
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
$backupRoot  = Join-Path $desktopPath "RevitServer_RVT_Backup"
$backupDate  = Join-Path $backupRoot  $dateStamp
$backupDest  = Join-Path $backupDate  "${version}_${serverHost}_${timeStamp}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null

Write-OK "Backup path:"
Write-Host "    $backupDest" -ForegroundColor White
Write-Host ""
Write-Host "  Tree preview:" -ForegroundColor DarkGray
Write-Host "  RevitServer_RVT_Backup\"                         -ForegroundColor DarkCyan
Write-Host "  └── $dateStamp\"                                  -ForegroundColor DarkCyan
Write-Host "      └── ${version}_${serverHost}_${timeStamp}\"  -ForegroundColor DarkCyan
foreach ($m in $models) {
    $displayPath = if ($m.SubFolder) { "          ├── $($m.SubFolder)\" } else { "" }
    if ($displayPath) { Write-Host $displayPath -ForegroundColor DarkGray }
    Write-Host "          │   └── $($m.Name)" -ForegroundColor DarkGray
}

# ─────────────────────────────────────────────────────────────
# STEP 8 — Export RVT files
# ─────────────────────────────────────────────────────────────
Write-Title "Step 8: Exporting RVT Files"

Write-Host "  Command: revitservertool.exe createLocalRVT" -ForegroundColor White
Write-Host "  Models do not need to be locked — tool reads live data." -ForegroundColor DarkGray
Write-Host ""

$successList  = [System.Collections.Generic.List[string]]::new()
$failList     = [System.Collections.Generic.List[string]]::new()
$current      = 0

foreach ($m in $models) {
    $current++
    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    # Mirror subfolder structure
    $targetDir = if ($m.SubFolder) {
        Join-Path $backupDest $m.SubFolder
    } else {
        $backupDest
    }

    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    # revitservertool.exe createLocalRVT "Path/Model.rvt" -s Hostname -d "DestFolder" -o
    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$targetDir`" -o"

    Write-Info "  -> $targetDir\$($m.Name)"

    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode

        if ($exitCode -eq 0) {
            Write-OK "Exported: $($m.Name)"
            $successList.Add($m.RSNPath)
        } else {
            Write-Fail "Failed (exit $exitCode): $($m.Name)"
            Write-Warn "  Common causes: model locked by user, insufficient permissions, corrupt model"
            $failList.Add("$($m.RSNPath)  [exit $exitCode]")
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
$manifestPath = Join-Path $backupDest "_BACKUP_MANIFEST.txt"
$manifest     = [System.Collections.Generic.List[string]]::new()
$manifest.Add("Revit Server RVT Backup Manifest")
$manifest.Add("=" * 64)
$manifest.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$manifest.Add("Revit Version : $version")
$manifest.Add("Server Host   : $serverHost")
$manifest.Add("Machine       : $($env:COMPUTERNAME)")
$manifest.Add("Tool          : $toolExe")
$manifest.Add("Projects Root : $projRoot")
$manifest.Add("Destination   : $backupDest")
$manifest.Add("Models total  : $modelCount")
$manifest.Add("Succeeded     : $([int]($successList | Measure-Object).Count)")
$manifest.Add("Failed        : $([int]($failList    | Measure-Object).Count)")
$manifest.Add("")

if ($successList.Count -gt 0) {
    $manifest.Add("SUCCEEDED:")
    $manifest.Add("-" * 64)
    foreach ($s in $successList) { $manifest.Add("  [OK]  $s") }
    $manifest.Add("")
}

if ($failList.Count -gt 0) {
    $manifest.Add("FAILED:")
    $manifest.Add("-" * 64)
    foreach ($f in $failList) { $manifest.Add("  [XX] $f") }
    $manifest.Add("")
}

$manifest | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest written: $manifestPath"

# ─────────────────────────────────────────────────────────────
# DONE — Summary
# ─────────────────────────────────────────────────────────────
$successCount = [int]($successList | Measure-Object).Count
$failCount    = [int]($failList    | Measure-Object).Count

Write-Title "Backup Complete"

Write-Host "  Server       : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Total models : $modelCount"  -ForegroundColor White
Write-Host "  Succeeded    : $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "  Failed       : $failCount"  -ForegroundColor Red
} else {
    Write-Host "  Failed       : 0"           -ForegroundColor Green
}
Write-Host ""
Write-Host "  Backup location:" -ForegroundColor White
Write-Host "  $backupDest"      -ForegroundColor Cyan
Write-Host ""
Write-Host "  Structure:"                                               -ForegroundColor DarkGray
Write-Host "  Desktop\RevitServer_RVT_Backup\"                         -ForegroundColor DarkCyan
Write-Host "  └── $dateStamp\"                                          -ForegroundColor DarkCyan
Write-Host "      └── ${version}_${serverHost}_${timeStamp}\"          -ForegroundColor DarkCyan
Write-Host "          ├── _BACKUP_MANIFEST.txt"                        -ForegroundColor DarkGray
Write-Host "          └── [mirrored tree of real .rvt files]"          -ForegroundColor DarkGray
Write-Host ""

$open = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($open.Trim().ToUpper() -eq "Y") {
    Start-Process explorer.exe $backupDest
}

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
