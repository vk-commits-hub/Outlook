# ============================================================
#  AI Reply Assistant — Outlook Sideloader for Windows
#  Double-click this file (or right-click > Run with PowerShell)
# ============================================================

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  AI Reply Assistant — Outlook Add-in Installer" -ForegroundColor Cyan
Write-Host "  -----------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

# ── 1. Find the manifest ──────────────────────────────────────────
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestSrc = Join-Path $scriptDir "manifest.xml"

if (-not (Test-Path $manifestSrc)) {
    Write-Host "  ERROR: manifest.xml not found next to this script." -ForegroundColor Red
    Write-Host "  Make sure all files are in the same folder." -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ── 2. Check the manifest has been configured ────────────────────
$manifestContent = Get-Content $manifestSrc -Raw
if ($manifestContent -match "YOUR_GITHUB_PAGES_URL") {
    Write-Host "  SETUP REQUIRED" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  You must set your hosting URL in manifest.xml before installing." -ForegroundColor Yellow
    Write-Host "  See README.md for how to host for free on GitHub Pages." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Quick option — if you just want to test locally:" -ForegroundColor Cyan
    Write-Host "  Run 'start-local-server.bat' first, then re-run this installer." -ForegroundColor Cyan
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ── 3. Copy manifest to the Outlook WEF trusted folder ───────────
$wefPath = "$env:APPDATA\Microsoft\Outlook\Wef"

if (-not (Test-Path $wefPath)) {
    New-Item -ItemType Directory -Path $wefPath -Force | Out-Null
    Write-Host "  Created Outlook add-in folder: $wefPath" -ForegroundColor DarkGray
}

$destManifest = Join-Path $wefPath "AIReplyAssistant.xml"
Copy-Item -Path $manifestSrc -Destination $destManifest -Force

Write-Host "  Manifest installed to:" -ForegroundColor Green
Write-Host "  $destManifest" -ForegroundColor DarkGray
Write-Host ""

# ── 4. Prompt to restart Outlook ────────────────────────────────
$outlook = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($outlook) {
    $restart = Read-Host "  Outlook is running. Close and restart it now? [Y/n]"
    if ($restart -ne 'n' -and $restart -ne 'N') {
        Write-Host "  Closing Outlook…" -ForegroundColor DarkGray
        $outlook | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "  Starting Outlook…" -ForegroundColor DarkGray
        Start-Process "outlook.exe"
    }
} else {
    Write-Host "  Outlook is not running. Start it to see the add-in." -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  Done! Look for the '✦ AI Reply Draft' button in the" -ForegroundColor Green
Write-Host "  Home ribbon when you open any email." -ForegroundColor Green
Write-Host ""
Read-Host "  Press Enter to exit"
