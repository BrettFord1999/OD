<#  OD-Migrate-Old.ps1
    Purpose: Prep and publish a safe migration “package” for Open Dental.

    New in this build:
    - Detects DB engine/version from the running service binary
    - Writes od_migration.json with engine/version and the proper Open Dental Trial link
    - Temporary SMB share allows write so NEW server can drop COPIED.txt
    - Grants NTFS Modify ONLY on the specific Package_* folder

    Run as Administrator on the OLD server.
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

function Get-ServiceBinaryPath {
  param([string]$ServiceName)
  $svcWmi = Get-WmiObject Win32_Service -Filter "Name='$ServiceName'" -ErrorAction SilentlyContinue
  if (-not $svcWmi) { return $null }
  # Extract the first quoted path or first token
  $path = $svcWmi.PathName
  if ($path -match '^"([^"]+)"') { return $matches[1] }
  else { return ($path.Split(' '))[0] }
}

function Detect-DBVersion {
  # Returns [pscustomobject] @{ Engine='MariaDB'|'MySQL'; VersionFull='10.5.22'; MajorMinor='10.5'; Bin='<path>' }
  $svc = Find-DbService
  if (-not $svc) { return $null }
  $bin = Get-ServiceBinaryPath -ServiceName $svc.Name
  if (-not (Test-Path $bin)) {
    Write-Warning "DB service found ($($svc.Name)) but binary path not found. Falling back to generic mapping."
    return [pscustomobject]@{ Engine=$svc.DisplayName; VersionFull=$null; MajorMinor=$null; Bin=$null }
  }
  $verOut = & $bin --version 2>&1
  $engine = ($verOut -match 'MariaDB') ? 'MariaDB' : 'MySQL'
  if ($verOut -match '(\d+\.\d+(\.\d+)?)') { $vf = $matches[1] } else { $vf = $null }
  $mm = $null; if ($vf) { $mm = ($vf -split '\.')[0..1] -join '.' }
  [pscustomobject]@{ Engine=$engine; VersionFull=$vf; MajorMinor=$mm; Bin=$bin }
}

function Get-ODTrialInfo {
  param([string]$Engine,[string]$MajorMinor)
  # Current Trial links from Open Dental (may change over time):
  $trialDefault = 'https://opendental.com/TrialDownload-24-3-54.exe'   # general trial
  $maria105     = 'https://opendental.com/TrialDownload-24-1-66.exe'   # MariaDB 10.5 trial
  $mysql55      = 'https://opendental.com/TrialDownload-20-5-63.exe'   # MySQL 5.5 trial
  $upgrade56    = 'https://opendental.com/site/mysql56update.html'     # MySQL 5.6 upgrade guide

  $trialUrl = $trialDefault
  $notes = @()

  if ($Engine -match 'Maria' -and $MajorMinor -eq '10.5') {
    $trialUrl = $maria105
    $notes += 'MariaDB 10.5 installer.'
  } elseif ($Engine -match 'MySQL' -and $MajorMinor -eq '5.5') {
    $trialUrl = $mysql55
    $notes += 'MySQL 5.5 installer.'
  } elseif ($Engine -match 'MySQL' -and $MajorMinor -eq '5.6') {
    # Per OD docs: install MySQL 5.5 trial first, then upgrade to 5.6 on the NEW server.
    $trialUrl = $mysql55
    $notes += 'Install MySQL 5.5 trial, then upgrade to 5.6 on NEW server.'
  } else {
    $notes += 'General trial installer (contains MariaDB).'
  }

  [pscustomobject]@{
    trial_url   = $trialUrl
    upgrade_url = ($Engine -match 'MySQL' -and $MajorMinor -eq '5.6') ? $upgrade56 : $null
    notes       = ($notes -join ' ')
  }
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
      New-SmbShare -Name $Name -Path $Path -ChangeAccess 'Everyone' -CachingMode None | Out-Null
    } else {
      try { Revoke-SmbShareAccess -Name $Name -AccountName 'Everyone' -Force -ErrorAction SilentlyContinue } catch {}
      Grant-SmbShareAccess -Name $Name -AccountName 'Everyone' -AccessRight Change -Force | Out-Null
    }
  } catch {
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
  try { icacls "$Path" /grant *S-1-1-0:(OI)(CI)M /T | Out-Null }
  catch {
    $acl  = Get-Acl $Path
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('Everyone','Modify','ContainerInherit, ObjectInherit','None','Allow')
    $acl.SetAccessRule($rule); Set-Acl -AclObject $acl -Path $Path
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

# Discover paths + DB version
$datadir = Guess-Data
