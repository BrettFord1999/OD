<#  OD-Migrate-New.ps1  (interactive, no creds, C:\OpenDentImages)
    Purpose: Pull Open Dental migration package from OLD server, read od_migration.json,
             download the right Trial installer if needed, restore DB & A-to-Z, bring NEW server online.

    Run as Administrator on the NEW server.
#>

[CmdletBinding()]
param(
  [string]$LocalStage = 'C:\OD_Migration\Imported',
  [switch]$SkipOpen3306
)

# --------- Helpers ---------
function Assert-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) { throw "Run this script as Administrator." }
}

function Find-DbService {
  Get-Service |
    Where-Object { $_.Name -match '^(mysql|mariadb)' -or $_.DisplayName -match '(MySQL|MariaDB)' } |
    Sort-Object Status -Descending |
    Select-Object -First 1
}

function Guess-DataDir {
  $cands = @(
    'C:\mysql\data',
    'C:\Program Files\MariaDB*\data',
    'C:\Program Files\MySQL\MySQL Server *\data',
    'C:\ProgramData\MySQL\MySQL Server *\Data',
    'C:\Program Files (x86)\MySQL\MySQL Server *\data'
  )
  $hit = $cands | ForEach-Object { Get-ChildItem -Path $_ -Directory -ErrorAction SilentlyContinue } | Select-Object -First 1
  if ($hit) { return $hit.FullName }
  'C:\mysql\data'
}

function Start-SafeCopy {
  param([string]$Source,[string]$Dest,[string]$LogPath)
  New-Item -ItemType Directory -Path $Dest -Force | Out-Null
  $args = @("$Source","$Dest","/MIR","/R:2","/W:2","/ZB","/COPY:DAT","/DCOPY:DAT","/FFT","/NP","/TEE","/LOG:$LogPath")
  & robocopy @args | Out-Null
  if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with code $LASTEXITCODE copying $Source" }
}

function Ensure-AtoZShare {
  param([string]$Path = 'C:\OpenDentImages')
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
  try {
    if (-not (Get-SmbShare -Name 'OpenDentImages' -ErrorAction SilentlyContinue)) {
      New-SmbShare -Name 'OpenDentImages' -Path $Path -FullAccess 'Everyone' -CachingMode None | Out-Null
    }
  } catch {
    cmd /c "net share OpenDentImages=`"$Path`" /GRANT:Everyone,FULL" | Out-Null
  }
}

function Prompt-Required([string]$message, [string]$default='') {
  $val = Read-Host "$message$(if($default){' ['+$default+']'})"
  if (-not $val -and $default) { $val = $default }
  if (-not $val) { throw "A value is required: $message" }
  $val
}

function Get-FileNameFromUrl([string]$url, [string]$fallback='OpenDentalTrial.exe') {
  try {
    $u = New-Object System.Uri($url)
    $name = [System.IO.Path]::GetFileName($u.AbsolutePath)
    if ([string]::IsNullOrWhiteSpace($name)) { return $fallback }
    return $name
  } catch { return $fallback }
}

# --------- Main ---------
Set-StrictMode -Version Latest
Assert-Admin

$SourceHost = Prompt-Required "Enter OLD server hostname or IP"
$ShareName  = Prompt-Required "Enter migration share name" "ODMigration"
$srcRoot    = "\\$SourceHost\$ShareName"

Write-Host "Connecting to: $srcRoot" -ForegroundColor Cyan
if (-not (Test-Path $srcRoot)) { throw "Cannot access $srcRoot. Ensure the OLD script created the share and firewall allows SMB." }

# Pick the latest Package_* folder
$pkg = Get-ChildItem -Path $srcRoot -Directory -ErrorAction SilentlyContinue |
       Where-Object { $_.Name -like 'Package_*' } |
       Sort-Object Name -Descending | Select-Object -First 1
if (-not $pkg) { throw "No 'Package_*' folder found under $srcRoot" }
$pkgPath = (Join-Path $srcRoot $pkg.Name)
Write-Host "Using package: $pkgPath" -ForegroundColor Cyan

# Copy package locally
New-Item -ItemType Directory -Force -Path $LocalStage | Out-Null
$localPkg = Join-Path $LocalStage $pkg.Name
Start-SafeCopy -Source $pkgPath -Dest $localPkg -LogPath (Join-Path $LocalStage 'import_pkg.log')

# Read metadata
$metaPath = Join-Path $localPkg 'od_migration.json'
$meta = $null
if (Test-Path $metaPath) {
  $meta = Get-Content $metaPath -Raw | ConvertFrom-Json
  Write-Host "OLD server DB: $($meta.engine) $($meta.version_full)  → Suggested Trial: $($meta.od_trial_url)" -ForegroundColor Gray
} else {
  Write-Warning "No od_migration.json found. Proceeding without auto-matched installer."
}

# Ensure a DB service exists; if not, download & run the suggested Trial installer
$svcDB = Find-DbService
if (-not $svcDB -and $meta -and $meta.od_trial_url) {
  $ans = Read-Host "DB service not found. Download and run the suggested Trial installer now? [Y/n]"
  if ($ans -notmatch '^(n|no)$') {
    $dlDir = Join-Path $LocalStage 'Installers'
    New-Item -ItemType Directory -Force -Path $dlDir | Out-Null
    $fileName = Get-FileNameFromUrl -url $meta.od_trial_url
    $out = Join-Path $dlDir $fileName
    Write-Host "Downloading $($meta.od_trial_url) …"
    try {
      Invoke-WebRequest -Uri $meta.od_trial_url -OutFile $out -UseBasicParsing
      Write-Host "Downloaded: $out" -ForegroundColor Green
      Write-Host "Launching installer (check only DB + grant tables + my.ini + OpenDentImages) …" -ForegroundColor Yellow
      Start-Process -FilePath $out -Verb RunAs -Wait
    } catch {
      Write-Warning "Download or install failed: $_"
      Write-Host "If needed, open in browser: $($meta.od_trial_url)"
      if ($meta.od_upgrade_url) { Write-Host "Upgrade guide: $($meta.od_upgrade_url)" }
    }
    # Re-detect service after install attempt
    $svcDB = Find-DbService
  }
}

if (-not $svcDB) {
  Write-Error "No MySQL/MariaDB service detected on NEW server. Install the Trial DB components first, then re-run this script."
  exit 1
}

# Stop DB to restore
if ($svcDB.Status -ne 'Stopped') { Stop-Service -Name $svcDB.Name -Force }

# Data dir backup + restore
$datadir   = Guess-DataDir
$backupDir = "$datadir`_before_$((Get-Date).ToString('yyyyMMdd_HHmmss'))"
if (Test-Path $datadir) {
  Write-Host "Backing up current DB to: $backupDir"
  Start-SafeCopy -Source $datadir -Dest $backupDir -LogPath (Join-Path $LocalStage 'backup_current_db.log')
}
Write-Host "Restoring DB from package\Database to: $datadir"
Start-SafeCopy -Source (Join-Path $localPkg 'Database') -Dest $datadir -LogPath (Join-Path $LocalStage 'restore_db.log')

# Start DB
Start-Service -Name $svcDB.Name

# A-to-Z: ALWAYS use C:\OpenDentImages
$AtoZPath = 'C:\OpenDentImages'
Ensure-AtoZShare -Path $AtoZPath
Write-Host "Restoring A-to-Z from package\OpenDentImages …"
Start-SafeCopy -Source (Join-Path $localPkg 'OpenDentImages') -Dest $AtoZPath -LogPath (Join-Path $LocalStage 'restore_atoz.log')

# Tell operator where Setup.exe is (if present)
$setupRoot  = Join-Path $AtoZPath 'Setup.exe'
$setupFiles = Join-Path $AtoZPath 'SetupFiles\Setup.exe'
if (Test-Path $setupRoot) {
  Write-Host "Run installer: $setupRoot  (Right-click → Run as administrator)" -ForegroundColor Green
} elseif (Test-Path $setupFiles) {
  Write-Host "Run installer: $setupFiles  (Right-click → Run as administrator)" -ForegroundColor Green
} else {
  Write-Warning "No Setup.exe found in $AtoZPath (or \SetupFiles). If needed, use the Trial installer to seed OpenDentImages with Setup.exe."
}

# Open inbound 3306 (optional)
if (-not $SkipOpen3306) {
  if (-not (Get-NetFirewallRule -DisplayName 'Allow MySQL 3306 (OpenDental)' -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName 'Allow MySQL 3306 (OpenDental)' -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3306 | Out-Null
    Write-Host "Opened inbound port 3306 on NEW server." -ForegroundColor Yellow
  }
}

# Try to signal OLD server (OLD script allows write)
$localSignal = Join-Path $localPkg 'COPIED.txt'
'copied=1' | Set-Content $localSignal -Encoding UTF8
try {
  $remoteSignal = Join-Path $pkgPath 'COPIED.txt'
  Copy-Item $localSignal -Destination $remoteSignal -Force
  Write-Host "Signaled OLD server by creating COPIED.txt." -ForegroundColor Green
} catch {
  Write-Warning "Could not write COPIED.txt back to the OLD share."
}

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Green
Write-Host "1) If not already installed, run Setup.exe from C:\OpenDentImages (or install via the Trial you just downloaded)."
Write-Host "2) Launch Open Dental (Run as admin) and point to THIS server."
Write-Host "3) Setup → Data Paths & Preferences → set A-to-Z to C:\OpenDentImages and Update Server to this machine."
Write-Host "4) Install/start OpenDentalService and eConnector if required."
if ($meta -and $meta.od_upgrade_url) {
  Write-Host "MySQL 5.6 note: follow upgrade guide → $($meta.od_upgrade_url)"
}
