#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Revit Server RVT Backup Script v5
    Works in two modes:
      - LOCAL  : run directly on the Revit Server host machine
      - REMOTE : run from any PC with network access to a Revit Server

.NOTES
    - Run AS ADMINISTRATOR
    - Supported versions: 2020-2027
    - Output: Desktop\RevitServer_RVT_Backup\20260406_1656_VERSION_HOST\
#>

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
function Write-Title {
    param([string]$Text)
    $line = "=" * 64
    Write-Host ""; Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text"            -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan; Write-Host ""
}
function Write-OK   { param([string]$M) Write-Host "  [OK]  $M" -ForegroundColor Green      }
function Write-Info { param([string]$M) Write-Host "  [..] $M"  -ForegroundColor DarkGray   }
function Write-Warn { param([string]$M) Write-Host "  [!!] $M"  -ForegroundColor Yellow     }
function Write-Fail { param([string]$M) Write-Host "  [XX] $M"  -ForegroundColor Red        }
function Write-Skip { param([string]$M) Write-Host "  [--] $M"  -ForegroundColor DarkYellow }

# ─────────────────────────────────────────────────────────────
# STEP 1 — Choose run mode: LOCAL or REMOTE
# ─────────────────────────────────────────────────────────────
Write-Title "Revit Server RVT Backup v5 — $($env:COMPUTERNAME)"

Write-Host "  Run mode:" -ForegroundColor White
Write-Host "    [1]  LOCAL  — this machine IS the Revit Server host" -ForegroundColor Cyan
Write-Host "    [2]  REMOTE — connect to a Revit Server over the network" -ForegroundColor Cyan
Write-Host ""
$modeInput = Read-Host "  Enter 1 or 2"

$isRemote = $false
switch ($modeInput.Trim()) {
    "1" { $isRemote = $false; Write-OK "Mode: LOCAL" }
    "2" { $isRemote = $true;  Write-OK "Mode: REMOTE" }
    default {
        Write-Fail "Invalid choice. Exiting."
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────
# STEP 2 — Determine server hostname
# ─────────────────────────────────────────────────────────────
Write-Title "Step 2: Server Hostname"

$serverHost = $null

if ($isRemote) {
    Write-Host "  Enter the hostname or IP of the Revit Server machine." -ForegroundColor Yellow
    Write-Host "  Example: MYSERVER  or  192.168.1.50" -ForegroundColor DarkGray
    Write-Host ""
    $serverHost = (Read-Host "  Hostname / IP").Trim()
    if ([string]::IsNullOrWhiteSpace($serverHost)) {
        Write-Fail "No hostname entered. Exiting."
        exit 1
    }
    Write-OK "Remote server: $serverHost"

    # Quick connectivity check
    Write-Info "Testing connectivity to $serverHost ..."
    $ping = Test-Connection -ComputerName $serverHost -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($ping) {
        Write-OK "Host is reachable."
    } else {
        Write-Warn "Host did not respond to ping — continuing anyway (ping may be blocked)."
    }
} else {
    $serverHost = $env:COMPUTERNAME
    Write-OK "Local machine: $serverHost"
}

# ─────────────────────────────────────────────────────────────
# STEP 3 — Choose Revit Server version
# ─────────────────────────────────────────────────────────────
Write-Title "Step 3: Revit Server Version"

$supportedVersions = 2020..2027

# Tool search paths — checked on THIS machine (where the tool must be installed)
$toolPathPatterns = @(
    "C:\Program Files\Autodesk\Autodesk Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit {VER}\tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe",
    "C:\Program Files (x86)\Autodesk\Revit Server {VER}\Tools\RevitServerToolCommand\revitservertool.exe"
)

# Projects folder search paths:
#   LOCAL  -> local disk paths
#   REMOTE -> UNC paths on the remote server
$projectsPathPatternsLocal = @(
    "C:\ProgramData\Autodesk\Revit Server {VER}\Projects",
    "C:\ProgramData\Autodesk\Autodesk Revit Server {VER}\Projects"
)
$projectsPathPatternsRemote = @(
    "\\{HOST}\C$\ProgramData\Autodesk\Revit Server {VER}\Projects",
    "\\{HOST}\C$\ProgramData\Autodesk\Autodesk Revit Server {VER}\Projects",
    "\\{HOST}\ProgramData\Autodesk\Revit Server {VER}\Projects"
)

$detectedVersions = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "  Scanning for Revit Server versions..." -ForegroundColor White
Write-Host ""

foreach ($v in $supportedVersions) {
    $toolExe  = $null
    $projRoot = $null

    # Tool is always searched on THIS local machine
    foreach ($pat in $toolPathPatterns) {
        $p = $pat.Replace("{VER}", "$v")
        if (Test-Path $p) { $toolExe = $p; break }
    }

    # Projects folder: local or remote UNC
    if ($isRemote) {
        foreach ($pat in $projectsPathPatternsRemote) {
            $p = $pat.Replace("{VER}", "$v").Replace("{HOST}", $serverHost)
            if (Test-Path $p -ErrorAction SilentlyContinue) { $projRoot = $p; break }
        }
    } else {
        foreach ($pat in $projectsPathPatternsLocal) {
            $p = $pat.Replace("{VER}", "$v")
            if (Test-Path $p) { $projRoot = $p; break }
        }
    }

    if ($toolExe -or $projRoot) {
        $detectedVersions.Add([PSCustomObject]@{
            Version      = "$v"
            ToolExe      = $toolExe
            ProjectsRoot = $projRoot
        })
        $tTag = if ($toolExe)  { "[TOOL OK]"     } else { "[NO TOOL — needed on this PC]" }
        $pTag = if ($projRoot) { "[PROJECTS OK]" } else { "[NO PROJECTS]" }
        Write-Host "  ► Revit Server $v  $tTag  $pTag" -ForegroundColor Green
    }
}

# If nothing auto-detected (especially in REMOTE mode), offer manual entry
$selectedEntry = $null

if ($detectedVersions.Count -eq 0) {
    Write-Warn "No Revit Server installation detected automatically."
    Write-Host ""
    Write-Host "  Enter the version manually (e.g. 2025):" -ForegroundColor Yellow
    $manualVer = (Read-Host "  Version").Trim()

    if ($manualVer -notin ($supportedVersions | ForEach-Object { "$_" })) {
        Write-Fail "Version '$manualVer' not in supported range (2020-2027). Exiting."
        exit 1
    }

    $selectedEntry = [PSCustomObject]@{
        Version      = $manualVer
        ToolExe      = $null
        ProjectsRoot = $null
    }
} elseif ($detectedVersions.Count -eq 1) {
    $selectedEntry = $detectedVersions[0]
    Write-OK "Auto-selected: Revit Server $($selectedEntry.Version)"
} else {
    Write-Host ""
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

$version  = $selectedEntry.Version
$toolExe  = $selectedEntry.ToolExe
$projRoot = $selectedEntry.ProjectsRoot

# ─────────────────────────────────────────────────────────────
# STEP 4 — Validate revitservertool.exe (must be on THIS machine)
# ─────────────────────────────────────────────────────────────
Write-Title "Step 4: Validating RevitServerTool"

if (-not $toolExe) {
    Write-Fail "revitservertool.exe NOT found for Revit $version on this machine."
    Write-Host ""
    if ($isRemote) {
        Write-Host "  REMOTE mode: the tool must be installed on THIS PC (the machine running the script)." -ForegroundColor Yellow
        Write-Host "  Install Revit $version or copy the full RevitServerToolCommand folder from the server." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Paths checked:" -ForegroundColor DarkGray
    foreach ($pat in $toolPathPatterns) {
        Write-Host "    $($pat.Replace('{VER}',$version))" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualTool = Read-Host "  Enter full path to revitservertool.exe (or Enter to abort)"
    if ([string]::IsNullOrWhiteSpace($manualTool) -or -not (Test-Path $manualTool.Trim())) {
        Write-Fail "Tool not found. Exiting."
        exit 1
    }
    $toolExe = $manualTool.Trim()
}

$toolFileVer = (Get-Item $toolExe).VersionInfo.FileVersion
Write-OK "Tool     : $toolExe"
Write-OK "Version  : $toolFileVer"

# ─────────────────────────────────────────────────────────────
# STEP 5 — Locate Projects folder
# ─────────────────────────────────────────────────────────────
Write-Title "Step 5: Locating Projects Folder"

if (-not $projRoot) {
    Write-Warn "Projects folder not auto-detected."
    Write-Host ""
    if ($isRemote) {
        Write-Host "  Expected UNC path: \\$serverHost\C$\ProgramData\Autodesk\Revit Server $version\Projects" -ForegroundColor DarkGray
    } else {
        Write-Host "  Expected path: C:\ProgramData\Autodesk\Revit Server $version\Projects" -ForegroundColor DarkGray
    }
    Write-Host ""
    $manualProj = Read-Host "  Enter full path to Projects folder"
    if (-not (Test-Path $manualProj.Trim() -ErrorAction SilentlyContinue)) {
        Write-Fail "Path not accessible. Exiting."
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

$models = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($f in $modelFolders) {
    $relPath = $f.FullName.Replace($projRoot, "").TrimStart("\")
    $rsnPath = $relPath.Replace("\", "/")
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
$stamp       = Get-Date -Format "yyyyMMdd_HHmm"
$backupDest  = Join-Path $desktopPath "RevitServer_RVT_Backup\${stamp}_${version}_${serverHost}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null

Write-OK "Backup folder:"
Write-Host "  $backupDest" -ForegroundColor White

# ─────────────────────────────────────────────────────────────
# STEP 8 — Export RVT files
# ─────────────────────────────────────────────────────────────
Write-Title "Step 8: Exporting RVT Files"

Write-Host "  Tool   : revitservertool.exe createLocalRVT" -ForegroundColor White
Write-Host "  Server : $serverHost  (Revit Server $version)" -ForegroundColor White
Write-Host "  Mode   : $(if ($isRemote) { 'REMOTE' } else { 'LOCAL' })" -ForegroundColor White
Write-Host "  Locked/busy models will be SKIPPED automatically." -ForegroundColor DarkGray
Write-Host ""

$successList = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$failList    = [System.Collections.Generic.List[string]]::new()
$current     = 0

foreach ($m in $models) {
    $current++
    Write-Host "  [$current/$modelCount] $($m.RSNPath)" -ForegroundColor White

    # Full destination file path — preserves Revit Server folder tree
    # e.g. backupDest\Viktoriia Ilytska\Viktoriia Ilytska.rvt
    $destFilePath = Join-Path $backupDest $m.RelPath
    $destFolder   = [System.IO.Path]::GetDirectoryName($destFilePath)

    if (-not (Test-Path $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    Write-Info "  -> $destFilePath"

    # revitservertool.exe createLocalRVT "RSN/Path.rvt" -s HOSTNAME -d "C:\full\path.rvt" -o
    $toolArgs = "createLocalRVT `"$($m.RSNPath)`" -s $serverHost -d `"$destFilePath`" -o"

    try {
        $proc     = Start-Process -FilePath $toolExe -ArgumentList $toolArgs -NoNewWindow -Wait -PassThru
        $exitCode = $proc.ExitCode

        switch ($exitCode) {
            0 {
                if (Test-Path $destFilePath) {
                    $sizeMB = [math]::Round((Get-Item $destFilePath).Length / 1MB, 1)
                    Write-OK "Exported: $($m.Name)  ($sizeMB MB)"
                    $successList.Add("$($m.RSNPath)  [$sizeMB MB]")
                } else {
                    Write-Warn "Exit 0 but file not found — may be a path issue: $destFilePath"
                    $failList.Add("$($m.RSNPath)  [exit 0 but file missing]")
                }
            }
            5 {
                Write-Skip "SKIPPED — locked by user: $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [exit 5 — locked]")
            }
            1 {
                Write-Skip "SKIPPED — model busy: $($m.Name)"
                $skipList.Add("$($m.RSNPath)  [exit 1 — busy]")
            }
            default {
                Write-Fail "FAILED (exit $exitCode): $($m.Name)"
                $failList.Add("$($m.RSNPath)  [exit $exitCode]")
            }
        }
    } catch {
        Write-Fail "Exception: $($_.Exception.Message)"
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
$manifest.Add("Revit Server RVT Backup Manifest v5")
$manifest.Add("=" * 64)
$manifest.Add("Date          : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$manifest.Add("Run Mode      : $(if ($isRemote) { 'REMOTE' } else { 'LOCAL' })")
$manifest.Add("Revit Version : $version")
$manifest.Add("Server Host   : $serverHost")
$manifest.Add("This Machine  : $($env:COMPUTERNAME)")
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
Write-OK "Manifest: $manifestPath"

# ─────────────────────────────────────────────────────────────
# DONE — Summary
# ─────────────────────────────────────────────────────────────
Write-Title "Backup Complete"

Write-Host "  Mode          : $(if ($isRemote) { 'REMOTE' } else { 'LOCAL' })"         -ForegroundColor White
Write-Host "  Server        : $serverHost  (Revit Server $version)"                     -ForegroundColor White
Write-Host "  Total models  : $modelCount"                                               -ForegroundColor White
Write-Host "  Succeeded     : $successCount"                                             -ForegroundColor Green
Write-Host "  Skipped       : $skipCount" -ForegroundColor $(if ($skipCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Failed        : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red"    } else { "Green" })
Write-Host ""
Write-Host "  Backup   : $backupDest"   -ForegroundColor Cyan
Write-Host "  Manifest : $manifestPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Folder structure:" -ForegroundColor DarkGray
Write-Host "  RevitServer_RVT_Backup\"                           -ForegroundColor DarkCyan
Write-Host "  └── ${stamp}_${version}_${serverHost}\"            -ForegroundColor DarkCyan
Write-Host "      ├── _BACKUP_MANIFEST.txt"                      -ForegroundColor DarkGray
foreach ($m in $models) {
    $parent = [System.IO.Path]::GetDirectoryName($m.RelPath)
    if ($parent) {
        Write-Host "      ├── $parent\"         -ForegroundColor DarkCyan
        Write-Host "      │   └── $($m.Name)" -ForegroundColor DarkGray
    } else {
        Write-Host "      ├── $($m.Name)"      -ForegroundColor DarkGray
    }
}
Write-Host ""

$open = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($open.Trim().ToUpper() -eq "Y") { Start-Process explorer.exe $backupDest }

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
