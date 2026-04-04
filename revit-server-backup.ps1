<#
.SYNOPSIS
    Revit Server Backup Script
    Reads the Revit Server Projects folder, mirrors the full folder tree
    to a dated backup folder on the Desktop, and copies all model files.

.NOTES
    - Run on the Revit Server HOST machine (or a machine with UNC access to it)
    - No admin rights required — only read access to the Projects share
    - All users must be out of the models before backing up (best practice)
    - Supported versions: 2020–2027
#>

# ─────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────
function Write-Title {
    param([string]$Text)
    $line = "=" * ($Text.Length + 4)
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}

function Write-OK   { param([string]$Msg) Write-Host "  [OK]  $Msg" -ForegroundColor Green }
function Write-Info { param([string]$Msg) Write-Host "  [..] $Msg"  -ForegroundColor DarkGray }
function Write-Warn { param([string]$Msg) Write-Host "  [!!] $Msg"  -ForegroundColor Yellow }
function Write-Fail { param([string]$Msg) Write-Host "  [XX] $Msg"  -ForegroundColor Red }

# ─────────────────────────────────────────────────────────────
# STEP 1 — Select Revit Server version
# ─────────────────────────────────────────────────────────────
Write-Title "Revit Server Backup Tool"

$supportedVersions = 2020..2027

Write-Host "  Available Revit Server versions:" -ForegroundColor White
foreach ($v in $supportedVersions) {
    $rsnPath           = "C:\ProgramData\Autodesk\Revit Server $v\Config\RSN.ini"
    $localProjectsPath = "C:\ProgramData\Autodesk\Revit Server $v\Projects"
    $exists = (Test-Path $rsnPath) -or (Test-Path $localProjectsPath)
    $tag    = if ($exists) { " [DETECTED]" } else { "" }
    Write-Host "    $v$tag" -ForegroundColor $(if ($exists) { "Green" } else { "DarkGray" })
}

Write-Host ""
$versionInput = Read-Host "  Enter Revit Server version (e.g. 2025)"
$version = $versionInput.Trim()

if ($version -notin ($supportedVersions | ForEach-Object { "$_" })) {
    Write-Fail "Version '$version' is not in the supported range (2020–2027). Exiting."
    exit 1
}
Write-OK "Selected version: Revit Server $version"

# ─────────────────────────────────────────────────────────────
# STEP 2 — Detect server host from RSN.ini
# ─────────────────────────────────────────────────────────────
Write-Title "Step 2: Detecting Server Host"

$rsnIniPath = "C:\ProgramData\Autodesk\Revit Server $version\Config\RSN.ini"
$serverHost = $null

if (Test-Path $rsnIniPath) {
    $rsnContent = Get-Content $rsnIniPath -ErrorAction SilentlyContinue
    $serverHost = ($rsnContent | Where-Object { $_ -match '\S' } | Select-Object -First 1).Trim()
    Write-OK "RSN.ini found at: $rsnIniPath"
    Write-OK "Server host from RSN.ini: $serverHost"
} else {
    Write-Warn "RSN.ini not found at: $rsnIniPath"
    Write-Host ""
    Write-Host "  Options:" -ForegroundColor White
    Write-Host "    1. Enter a server hostname or IP manually"
    Write-Host "    2. Type LOCAL  (this machine is the Revit Server host)"
    Write-Host ""
    $hostInput = Read-Host "  Enter hostname/IP or type LOCAL"
    if ($hostInput.Trim().ToUpper() -eq "LOCAL") {
        $serverHost = $env:COMPUTERNAME
        Write-OK "Using local machine as host: $serverHost"
    } else {
        $serverHost = $hostInput.Trim()
        Write-OK "Using entered host: $serverHost"
    }
}

# ─────────────────────────────────────────────────────────────
# STEP 3 — Determine Projects folder path
# ─────────────────────────────────────────────────────────────
Write-Title "Step 3: Locating Projects Folder"

$candidatePaths = @(
    "C:\ProgramData\Autodesk\Revit Server $version\Projects",
    "\\$serverHost\Autodesk`$\Revit Server $version\Projects",
    "\\$serverHost\ProgramData\Autodesk\Revit Server $version\Projects",
    "\\$serverHost\C$\ProgramData\Autodesk\Revit Server $version\Projects"
)

$projectsRoot = $null
foreach ($path in $candidatePaths) {
    Write-Info "Checking: $path"
    if (Test-Path $path -ErrorAction SilentlyContinue) {
        $projectsRoot = $path
        Write-OK "Projects folder found: $projectsRoot"
        break
    }
}

if (-not $projectsRoot) {
    Write-Fail "Could not locate the Projects folder automatically."
    Write-Host ""
    Write-Host "  Please enter the full path to the Projects folder manually:" -ForegroundColor Yellow
    Write-Host "  Example: \\\\myserver\\Autodesk`$\\Revit Server $version\\Projects" -ForegroundColor DarkGray
    Write-Host "  Example: C:\ProgramData\Autodesk\Revit Server $version\Projects"    -ForegroundColor DarkGray
    Write-Host ""
    $manualPath = Read-Host "  Projects folder path"
    if (-not (Test-Path $manualPath.Trim())) {
        Write-Fail "Path does not exist or is not accessible. Exiting."
        exit 1
    }
    $projectsRoot = $manualPath.Trim()
    Write-OK "Using manually entered path: $projectsRoot"
}

# ─────────────────────────────────────────────────────────────
# STEP 4 — Scan and display the model tree
# ─────────────────────────────────────────────────────────────
Write-Title "Step 4: Scanning Revit Server Model Tree"

$allItems     = Get-ChildItem -Path $projectsRoot -Recurse -ErrorAction SilentlyContinue
$modelFolders = @($allItems | Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" })
$orgFolders   = @($allItems | Where-Object { $_.PSIsContainer -and $_.Name -notlike "*.rvt" })

$modelCount = [int]($modelFolders | Measure-Object).Count
$orgCount   = [int]($orgFolders   | Measure-Object).Count

Write-OK "Organisational folders : $orgCount"
Write-OK "Model folders (.rvt)   : $modelCount"
Write-Host ""

if ($modelCount -eq 0) {
    Write-Warn "No .rvt model folders found under: $projectsRoot"
    Write-Warn "Folder may be empty or you may lack read permissions."
    $continue = Read-Host "  Continue anyway and back up all files? (Y/N)"
    if ($continue.Trim().ToUpper() -ne "Y") { exit 0 }
}

Write-Host "  Model tree preview:" -ForegroundColor White
foreach ($m in $modelFolders) {
    $rel = $m.FullName.Replace($projectsRoot, "").TrimStart("\")
    Write-Host "    RSN://$serverHost/$rel" -ForegroundColor DarkCyan
}

# ─────────────────────────────────────────────────────────────
# STEP 5 — Create dated backup destination on Desktop
# ─────────────────────────────────────────────────────────────
Write-Title "Step 5: Creating Backup Destination"

$desktopPath = [Environment]::GetFolderPath("Desktop")
$dateStamp   = Get-Date -Format "yyyy-MM-dd"
$timeStamp   = Get-Date -Format "HH-mm"
$backupRoot  = Join-Path $desktopPath "RevitServer_Backup"
$backupDate  = Join-Path $backupRoot  "$dateStamp"
$backupDest  = Join-Path $backupDate  "RevitServer_${version}_${serverHost}_${timeStamp}"

New-Item -ItemType Directory -Path $backupDest -Force | Out-Null
Write-OK "Backup destination created:"
Write-Host "    $backupDest" -ForegroundColor White

# ─────────────────────────────────────────────────────────────
# STEP 6 — Write manifest file (FIXED)
# ─────────────────────────────────────────────────────────────
$manifestPath  = Join-Path $backupDest "_BACKUP_MANIFEST.txt"

$manifestLines = [System.Collections.Generic.List[string]]::new()
$manifestLines.Add("Revit Server Backup Manifest")
$manifestLines.Add("=" * 60)
$manifestLines.Add("Date         : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$manifestLines.Add("Revit Version: $version")
$manifestLines.Add("Server Host  : $serverHost")
$manifestLines.Add("Source       : $projectsRoot")
$manifestLines.Add("Destination  : $backupDest")
$manifestLines.Add("Models found : $modelCount")
$manifestLines.Add("")
$manifestLines.Add("Model List (RSN paths):")
$manifestLines.Add("-" * 60)

foreach ($m in $modelFolders) {
    $rel = $m.FullName.Replace($projectsRoot, "").TrimStart("\").Replace("\", "/")
    $manifestLines.Add("  RSN://$serverHost/$rel")
}

$manifestLines | Out-File -FilePath $manifestPath -Encoding UTF8
Write-OK "Manifest written: $manifestPath"

# ─────────────────────────────────────────────────────────────
# STEP 7 — Robocopy full tree
# ─────────────────────────────────────────────────────────────
Write-Title "Step 6: Copying Files (Robocopy)"

Write-Host "  This may take a while depending on model sizes..." -ForegroundColor Yellow
Write-Host ""

$robocopyLog  = Join-Path $backupDest "_robocopy.log"
$robocopyArgs = @(
    "`"$projectsRoot`"",
    "`"$backupDest`"",
    "/E",
    "/COPYALL",
    "/R:2",
    "/W:5",
    "/MT:8",
    "/NP",
    "/TEE",
    "/LOG:`"$robocopyLog`""
)

Write-Info "Running robocopy..."
$proc     = Start-Process -FilePath "robocopy" -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru
$exitCode = $proc.ExitCode

if ($exitCode -le 7) {
    Write-OK "Robocopy completed successfully (exit code: $exitCode)"
} else {
    Write-Fail "Robocopy encountered errors (exit code: $exitCode)"
    Write-Warn "Check the log: $robocopyLog"
}

# ─────────────────────────────────────────────────────────────
# STEP 8 — Summary
# ─────────────────────────────────────────────────────────────
Write-Title "Backup Complete"

$copiedItems  = Get-ChildItem -Path $backupDest -Recurse -ErrorAction SilentlyContinue
$copiedModels = [int](@($copiedItems | Where-Object { $_.PSIsContainer -and $_.Name -like "*.rvt" }) | Measure-Object).Count
$totalSizeMB  = [math]::Round(
    ($copiedItems | Where-Object { -not $_.PSIsContainer } | Measure-Object -Property Length -Sum).Sum / 1MB, 2)

Write-Host "  Source            : $projectsRoot"    -ForegroundColor White
Write-Host "  Destination       : $backupDest"      -ForegroundColor White
Write-Host "  Models backed up  : $copiedModels"    -ForegroundColor Green
Write-Host "  Total size        : $totalSizeMB MB"  -ForegroundColor Green
Write-Host "  Manifest          : $manifestPath"    -ForegroundColor Cyan
Write-Host "  Robocopy log      : $robocopyLog"     -ForegroundColor Cyan
Write-Host ""
Write-Host "  RevitServer_Backup\"                                          -ForegroundColor DarkCyan
Write-Host "  └── $dateStamp\"                                              -ForegroundColor DarkCyan
Write-Host "      └── RevitServer_${version}_${serverHost}_${timeStamp}\"  -ForegroundColor DarkCyan
Write-Host "          ├── _BACKUP_MANIFEST.txt"                            -ForegroundColor DarkGray
Write-Host "          ├── _robocopy.log"                                   -ForegroundColor DarkGray
Write-Host "          └── [mirrored Projects tree]"                        -ForegroundColor DarkGray
Write-Host ""

$openExplorer = Read-Host "  Open backup folder in Explorer? (Y/N)"
if ($openExplorer.Trim().ToUpper() -eq "Y") {
    Start-Process explorer.exe $backupDest
}

Write-Host ""
Write-Host "  Done. Press Enter to exit." -ForegroundColor Green
Read-Host | Out-Null
