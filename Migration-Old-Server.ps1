<#  OD-Migrate-Old.ps1
    Purpose: Prep and publish a safe migration “package” for Open Dental.

    Changes in this version:
    - Temporary SMB share grants Everyone **Change** (write) so NEW server can write COPIED.txt
    - NTFS permission: grant Modify to Everyone on the **package folder only**
#>

[CmdletBinding()]
param(
  [string]$StagingRoot = 'C:\OD_Migration',
  [string]$AtoZPathHint = 'C:\OpenDentImages',
  [string]$ShareName = 'ODMigration',
  [switch]$SkipRenameAtoZ,
  [switch]$NoFirewallBlock3306,
  [switch]$ConfirmAtoZ
)

# ----------------- Helpers -----------------
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
  $hit = $cands |
    ForEach-Object { Get-Item -LiteralPath $_ -ErrorAction SilentlyContinue } |
    Where-Object   { $_ -and $_.PSIsContainer } |
    Select-Object  -First 1
  if ($hit) { return $hit.FullName }
  Write-Warning "Could not auto-detect MySQL/MariaDB data dir. Using C:\mysql\data as fallback."
  'C:\mysql\data'
}

function Guess-MyIni {
  $cands = @(
    'C:\ProgramData\MySQL\my.ini',
    'C:\ProgramData\MySQL\MySQL Server 5.5\my.ini',
    'C:\Program Files\MariaDB*\data\my.ini',
    'C:\mysql\my.ini',
    'C:\Windows\my.ini'
  )
  $hit = $cands |
    ForEach-Object { Get-Item -LiteralPath $_ -ErrorAction SilentlyContinue } |
    Select-Object  -First 1
  if ($hit) { $hit.FullName } else { $null }
}

# --- A-to-Z auto-discovery (OLD server) ---
function Test-AtoZStructure {
  param([string]$Path)
  if (-not (Test-Path $Path -PathType Container)) { return 0 }
  $letters = [char[]](65..90) | ForEach-Object { [string]$_ }
  $letterCount = ($letters | Where-Object { Test-Path (Join-Path $Path $_) -PathType Container }).Count
  $bonus = 0
  foreach ($n in 'Mounts','EmailAttachments','Forms','Imaging','Scans') {
    if (Test-Path (Join-Path $Path $n) -PathType Container) { $bonus += 2 }
  }
  $letterCount + $bonus
}

function Find-AtoZPath {
  param(
    [string]$Hint = 'C:\OpenDentImages',
    [switch]$ConfirmAtoZ,
    [int]$MaxDepth = 3
  )
  $candidates = @()

  try {
    Get-SmbShare -ErrorAction Stop |
      Where-Object { $_.Name -match 'OpenDentImages' -or $_.Path -match 'OpenDentImages' } |
      ForEach-Object {
        $candidates += [pscustomobject]@{ Path=$_.Path; Score=(Test-AtoZStructure $_.Path); Source='Share' }
      }
  } catch {}

  foreach ($p in @($Hint,'C:\OpenDentImages','D:\OpenDentImages','E:\OpenDentImages')) {
    if (Test-Path $p) { $candidates += [pscustomobject]@{ Path=$p; Score=(Test-AtoZStructure $p); Source='Common' } }
  }

  $roots = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -ne $null }
  foreach ($r in $roots) {
    $queue = @([pscustomobject]@{Path=$r.Root;Depth=0})
    while ($queue.Count -gt 0) {
      $cur = $queue[0]
      if ($queue.Count -gt 1) { $queue = $queue[1..($queue.Count-1)] } else { $queue = @() }
      if ($cur.Depth -gt $MaxDepth) { continue }
      $dirs = Get-ChildItem -LiteralPath $cur.Path -Directory -ErrorAction SilentlyContinue
      foreach ($d in $dirs) {
        if ($d.Name -like 'OpenDentImages*') {
          $candidates += [pscustomobject]@{ Path=$d.FullName; Score=(Test-AtoZStructure $d.FullName); Source='Scan' }
        }
        if ($cur.Depth -lt $MaxDepth) {
          $queue += [pscustomobject]@{Path=$d.FullName;Depth=($cur.Depth+1)}
        }
      }
    }
  }

  $pick = $candidates | Sort-Object Score -Descending | Select-Object -First 1
  if (-not $pick -or $pick.Score -lt 8) {
    $pick = [pscustomobject]@{ Path=$Hint; Score=0; Source='Fallback' }
  }

  if ($ConfirmAtoZ -and $candidates.Count -gt 1) {
    Write-Host "A-to-Z candidates (score):" -ForegroundColor Cyan
    $ordered = $candidates | Sort-Object Score -Descending
    for ($i=0; $i -lt $ordered.Count; $i++) {
      "{0}. {1}  (Score {2}, {3})" -f $i,$ordered[$i].Path,$ordered[$i].Score,$ordered[$i].Source | Write-Host
    }
    $sel = Read-Host "Pick index (default 0)"
    if ($sel -match '^\d+$' -and [int]$sel -lt $ordered.Count) { $pick = $ordered[[int]$sel] }
  }

  $pick.Path
}

function Start-SafeCopy {
  param([string]$Source,[string]$Dest,[string]$LogPath)
  New-Item -ItemType Directory -Path $Dest -Force | Out-Null
  $args = @("$Source","$Dest","/MIR","/R:2","/W:2","/ZB","/COPY:DAT","/DCOPY:DAT","/FFT","/NP","/TEE","/LOG:$LogPath")
  & robocopy @args | Out-Null
  if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with code $LASTEXITCODE copying $Source" }
}

function Ensure-Share {
  param([string]$Path,[string]$Name)
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
  try {
    $share = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
    if (-not $share) {
      # Share with CHANGE (write) for Everyone
      New-SmbShare -Name $Name -Path $Path -ChangeAccess 'Everyone' -CachingMode None | Out-Null
    } else {
      # Ensure Everyone has CHANGE
      try { Revoke-SmbShareAccess -Name $Name -AccountName 'Everyone' -Force -ErrorAction SilentlyContinue } catch {}
      Grant-SmbShareAccess -Name $Name -AccountName 'Everyone' -AccessRight Change -Force | Out-Null
    }
  } catch {
    # Fallback (legacy): grant CHANGE at share level
    cmd /c "net share $Name=`"$Path`" /GRANT:Everyone,CHANGE" | Out-Null
  }
}

function Remove-Share {
  param([string]$Name)
  try { Get-SmbShare -Name $Name -ErrorAction Stop | Remove-SmbShare -Force }
  catch { cmd /c "net share $Name /delete /y" | Out-Null }
}

function Grant-NTFSModifyToEveryone {
  param([string]$Path)
  if (-not (Test-Path $Path -PathType Container)) { return }
  try {
    # Use SID for Everyone to avoid localization issues
    icacls "$Path" /grant *S-1-1-0:(OI)(CI)M /T | Out-Null
  } catch {
    # Fallback using .NET ACLs
    $acl  = Get-Acl $Path
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('Everyone','Modify','ContainerInherit, ObjectInherit','None','Allow')
    $acl.SetAccessRule($rule)
    Set-Acl -AclObject $acl -Path $Path
  }
}

# ----------------- Main -----------------
Set-StrictMode -Version Latest
Assert-Admin

$timeTag = Get-Date -Format 'yyyyMMdd_HHmmss'
$pkgRoot = Join-Path $StagingRoot "Package_$timeTag"
$pkgDb   = Join-Path $pkgRoot  'Database'
$pkgIni  = Join-Path $pkgRoot  'my.ini'
$pkgImg  = Join-Path $pkgRoot  'OpenDentImages'
$logDir  = Join-Path $pkgRoot  'logs'
New-Item -ItemType Directory -Force -Path $pkgRoot,$logDir | Out-Null

# Stop & disable services
$svcDB   = Find-DbService
$svcOdds = @('OpenDentalService','OpenDentalEConnector','OpenDentalAPIService') |
           ForEach-Object { Get-Service -Name $_ -ErrorAction SilentlyContinue } | Where-Object { $_ }

Write-Verbose "Stopping OpenDental/MySQL services…"
$toStop = @($svcDB) + $svcOdds | Where-Object { $_ }
foreach ($s in $toStop) { if ($s.Status -ne 'Stopped') { Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue } }

Write-Verbose "Disabling services…"
foreach ($s in $toStop) { try { Set-Service -Name $s.Name -StartupType Disabled } catch {} }

# Discover paths
$datadir = Guess-DataDir
$myini   = Guess-MyIni
$atoz    = Find-AtoZPath -Hint $AtoZPathHint -ConfirmAtoZ:$ConfirmAtoZ

Write-Host "Database data dir : $datadir"
if ($myini) { Write-Host "my.ini            : $myini" } else { Write-Warning "my.ini not found; continuing." }
Write-Host "OpenDentImages    : $atoz"
Write-Host ""

# Copy DB data (cold copy)
Write-Verbose "Copying database data folder…"
Start-SafeCopy -Source $datadir -Dest $pkgDb -LogPath (Join-Path $logDir 'copy_db.log')

# Copy my.ini
if ($myini) { Copy-Item -Path $myini -Destination $pkgIni -Force }

# Copy A-to-Z
Write-Verbose "Copying OpenDentImages (A-to-Z)…"
Start-SafeCopy -Source $atoz -Dest $pkgImg -LogPath (Join-Path $logDir 'copy_atoz.log')

# Optionally rename A-to-Z to prevent writes
if (-not $SkipRenameAtoZ) {
  $parent = Split-Path $atoz -Parent
  $newName = "OpenDentImages_old_$timeTag"
  $newPath = Join-Path $parent $newName
  try {
    Rename-Item -Path $atoz -NewName $newName -ErrorAction Stop
    Write-Host "Renamed A-to-Z to: $newPath  (prevents accidental writes)" -ForegroundColor Yellow
  } catch {
    Write-Warning "Could not rename A-to-Z: $_"
  }
}

# Temp share for migration **with write enabled**
Ensure-Share -Path $StagingRoot -Name $ShareName

# Grant NTFS Modify to Everyone on the specific package folder only (so COPIED.txt can be created)
Grant-NTFSModifyToEveryone -Path $pkgRoot

$shareUNC = "\\$($env:COMPUTERNAME)\$ShareName"
"ready=$timeTag`nsource=$shareUNC`npackage=$(Split-Path $pkgRoot -Leaf)" | Set-Content (Join-Path $pkgRoot 'READY.txt') -Encoding UTF8
Write-Host "Temporary share exposed at: $shareUNC  (Everyone: Change)" -ForegroundColor Cyan
Write-Host "Package located in: $pkgRoot" -ForegroundColor Cyan

# Optionally block 3306 inbound on OLD server
if (-not $NoFirewallBlock3306) {
  if (-not (Get-NetFirewallRule -DisplayName 'Block MySQL 3306 (OD Migration)' -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName 'Block MySQL 3306 (OD Migration)' -Direction Inbound -Action Block -Protocol TCP -LocalPort 3306 | Out-Null
    Write-Host "Blocked inbound port 3306 on OLD server." -ForegroundColor Yellow
  }
}

# Wait for new server signal
Write-Host "`n=== Waiting for new server to write COPIED.txt … Ctrl+C to abort ==="
$copiedFlag = Join-Path $pkgRoot 'COPIED.txt'
while (-not (Test-Path $copiedFlag)) { Start-Sleep -Seconds 5 }

# Cleanup share
Write-Host "New server signaled copy complete." -ForegroundColor Green
Write-Host "Removing temporary share…" -ForegroundColor Gray
Remove-Share -Name $ShareName

Write-Host "Old server is locked down. Leave services disabled to prevent clients reconnecting." -ForegroundColor Yellow
