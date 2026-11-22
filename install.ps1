# 0x15k Office AI - Automated Installer for Windows
# Configures Node.js dependencies, SSL certificates, Windows Office Sideloading, and background service.

param(
    [switch]$SideloadOnly,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ManifestPath = Join-Path $ProjectDir "manifest.xml"
$Guid = "76dc8e20-d4be-457e-b8a0-354a30087981"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "0x15k Office AI for Microsoft Office - Windows Setup" -ForegroundColor Cyan
Write-Host "Author: 0x15k - YL" -ForegroundColor Gray
Write-Host "==========================================================" -ForegroundColor Cyan

if ($Uninstall) {
    Write-Host "[0x15k-office] Uninstalling from Windows..." -ForegroundColor Yellow
    try {
        Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name $Guid -ErrorAction SilentlyContinue
        Write-Host "[0x15k-office] Removed Office developer registry entry." -ForegroundColor Green
    } catch {}
    
    try {
        $pids = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique
        foreach ($p in $pids) {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    Write-Host "[0x15k-office] Uninstall complete." -ForegroundColor Green
    exit 0
}

# 1. Check Node.js
Write-Host "[0x15k-office] Checking Node.js installation..." -ForegroundColor Yellow
try {
    $nodeVer = & node -v
    Write-Host "[0x15k-office] Found Node.js: $nodeVer" -ForegroundColor Green
} catch {
    Write-Host "[0x15k-office] ERROR: Node.js is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Node.js from https://nodejs.org/ (LTS recommended) and rerun this script." -ForegroundColor Yellow
    exit 1
}

# 2. Check/Install npm dependencies
if (-not $SideloadOnly) {
    if (-not (Test-Path (Join-Path $ProjectDir "node_modules"))) {
        Write-Host "[0x15k-office] Installing dependencies (npm install)..." -ForegroundColor Yellow
        Push-Location $ProjectDir
        & npm install --no-fund --no-audit
        Pop-Location
    }
}

# 3. Sideload Manifest into Windows Office (Word, Excel, PowerPoint)
Write-Host "[0x15k-office] Configuring Windows Office Sideloading (WEF Developer Registry)..." -ForegroundColor Yellow
try {
    $wefDevKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
    if (-not (Test-Path $wefDevKey)) {
        New-Item -Path $wefDevKey -Force | Out-Null
    }
    Set-ItemProperty -Path $wefDevKey -Name $Guid -Value $ManifestPath -Force
    
    $wefKey = "HKCU:\Software\Microsoft\Office\16.0\WEF"
    if (-not (Test-Path $wefKey)) {
        New-Item -Path $wefKey -Force | Out-Null
    }
    Set-ItemProperty -Path $wefKey -Name "OfficeWebAddinDeveloperExtras" -Value 1 -Type DWord -Force
    
    Write-Host "[0x15k-office] Successfully registered manifest in Windows Office Developer catalog." -ForegroundColor Green
} catch {
    Write-Host "[0x15k-office] WARNING: Could not write developer registry key: $_" -ForegroundColor Yellow
}

# 4. Optional: Copy to LocalAppData Wef cache if present
$wefLocalCache = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
if ($env:LOCALAPPDATA -and (Test-Path $wefLocalCache)) {
    try {
        Copy-Item -Path $ManifestPath -Destination (Join-Path $wefLocalCache "manifest.xml") -Force
        Write-Host "[0x15k-office] Sideloaded manifest to local Office cache: $wefLocalCache" -ForegroundColor Green
    } catch {}
}

if ($SideloadOnly) {
    Write-Host "[0x15k-office] Sideload complete. Restart Word, Excel, or PowerPoint to view the add-in." -ForegroundColor Green
    exit 0
}

# 5. Start background server
Write-Host "[0x15k-office] Starting 0x15k Office AI HTTPS server in background..." -ForegroundColor Yellow
try {
    $existing = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "[0x15k-office] Server is already running on https://localhost:3000" -ForegroundColor Green
    } else {
        Start-Process -FilePath "node" -ArgumentList (Join-Path $ProjectDir "server.js") -WorkingDirectory $ProjectDir -WindowStyle Hidden
        Start-Sleep -Seconds 2
        Write-Host "[0x15k-office] Server started on https://localhost:3000" -ForegroundColor Green
    }
} catch {
    Write-Host "[0x15k-office] Note: You can manually start the server by running 'npm start'." -ForegroundColor Yellow
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "Setup Completed Successfully!" -ForegroundColor Green
Write-Host "1. Open Microsoft Word, Excel, or PowerPoint." -ForegroundColor White
Write-Host "2. On the 'Home' ribbon tab, click '0x15k Office'." -ForegroundColor White
Write-Host "   (Or go to Insert -> Add-ins -> My Add-ins -> Developer Add-ins)" -ForegroundColor Gray
Write-Host "==========================================================" -ForegroundColor Cyan
