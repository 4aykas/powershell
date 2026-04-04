<#
.SYNOPSIS
    Opens Autodesk trial download pages for Revit, AutoCAD, and Navisworks
    (last 3 versions: 2025, 2026, 2027) in the default browser.

.NOTES
    Autodesk requires sign-in to generate authenticated download links.
    This script opens all 9 trial pages so you can download each manually.
    Trials are 30 days full-featured.
#>

Write-Host ""
Write-Host "Opening Autodesk Trial download pages..." -ForegroundColor Cyan
Write-Host "Sign in with your Autodesk account on each page to start the download." -ForegroundColor Yellow
Write-Host ""

$trialPages = @(
    # ── REVIT ────────────────────────────────────────────────
    [PSCustomObject]@{ Product = "Revit 2027";      URL = "https://www.autodesk.com/products/revit/free-trial" },
    [PSCustomObject]@{ Product = "Revit 2026";      URL = "https://www.autodesk.com/products/revit/free-trial" },
    [PSCustomObject]@{ Product = "Revit 2025";      URL = "https://www.autodesk.com/products/revit/free-trial" },

    # ── AUTOCAD ──────────────────────────────────────────────
    [PSCustomObject]@{ Product = "AutoCAD 2027";    URL = "https://www.autodesk.com/products/autocad/free-trial" },
    [PSCustomObject]@{ Product = "AutoCAD 2026";    URL = "https://www.autodesk.com/products/autocad/free-trial" },
    [PSCustomObject]@{ Product = "AutoCAD 2025";    URL = "https://www.autodesk.com/products/autocad/free-trial" },

    # ── NAVISWORKS ───────────────────────────────────────────
    [PSCustomObject]@{ Product = "Navisworks 2027"; URL = "https://www.autodesk.com/products/navisworks/free-trial" },
    [PSCustomObject]@{ Product = "Navisworks 2026"; URL = "https://www.autodesk.com/products/navisworks/free-trial" },
    [PSCustomObject]@{ Product = "Navisworks 2025"; URL = "https://www.autodesk.com/products/navisworks/free-trial" }
)

foreach ($item in $trialPages) {
    Write-Host "  Opening: $($item.Product)" -ForegroundColor DarkGray
    Start-Process $item.URL
    Start-Sleep -Seconds 2   # slight delay so tabs open cleanly
}

Write-Host ""
Write-Host "All 9 trial pages opened in your browser." -ForegroundColor Green
Write-Host ""
Write-Host "On each page:" -ForegroundColor Cyan
Write-Host "  1. Click 'Download free trial'"
Write-Host "  2. Sign in (or create a free Autodesk account)"
Write-Host "  3. Select the correct VERSION and language"
Write-Host "  4. Download the installer"
Write-Host ""
Write-Host "NOTE: Each trial is valid for 30 days from first launch." -ForegroundColor Yellow
Write-Host "NOTE: AutoCAD trial is 15 days. Revit and Navisworks are 30 days." -ForegroundColor Yellow
