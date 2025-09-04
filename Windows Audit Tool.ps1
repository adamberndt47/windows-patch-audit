[CmdletBinding()]
param(
  [ValidateSet('CSV','JSON','HTML')]
  [string]$Format = 'CSV',
  [string]$OutputPath = ".\PatchAudit",
  [switch]$UsePSWindowsUpdate = $true,
  [switch]$MicrosoftUpdate
)

function Get-RegistryValue {
  param($Path,$Name)
  try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
}

function Get-WSUSServer {
  $path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
  $srv  = Get-RegistryValue -Path $path -Name 'WUServer'
  return $srv
}

function Get-RebootPending {
  $keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations'
  )
  foreach ($k in $keys) {
    if (Test-Path $k) { return $true }
  }
  return $false
}

function Get-WULastTimes {
  $auPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
  $scanPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect'
  [pscustomobject]@{
    LastInstallTime = Get-RegistryValue -Path $auPath -Name 'LastSuccessTime'
    LastScanTime    = Get-RegistryValue -Path $scanPath -Name 'LastSuccessTime'
  }
}

function Get-PendingUpdatesCOM {
  # Fallback via Windows Update Agent COM
  try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $crit = "IsInstalled=0 and Type='Software'"
    $result = $searcher.Search($crit)
    $pending = @()
    foreach ($upd in $result.Updates) {
      # severity not always available; attempt extraction from MsrcSeverity or Severity if present
      $sev = $null
      try { $sev = $upd.MsrcSeverity } catch {}
      try { if (-not $sev) { $sev = $upd.Severity } } catch {}
      $kbs = ($upd.KBArticleIDs -join ',')
      $pending += [pscustomobject]@{
        KB        = $kbs
        Title     = $upd.Title
        Severity  = $sev
      }
    }
    return $pending
  } catch {
    Write-Verbose "COM search failed: $($_.Exception.Message)"
    return @()
  }
}

function Get-PendingUpdatesPSWU {
  try {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) { return $null }
    Import-Module PSWindowsUpdate -ErrorAction Stop | Out-Null
    $params = @{ MicrosoftUpdate = $MicrosoftUpdate.IsPresent }
    $updates = Get-WindowsUpdate -Criteria "IsInstalled=0" -IsHidden:$false -IgnoreUserInput -ErrorAction Stop @params
    # Map to common shape
    $updates | ForEach-Object {
      [pscustomobject]@{
        KB       = ($_.KB -join ',')
        Title    = $_.Title
        Severity = $_.MsrcSeverity
      }
    }
  } catch {
    Write-Verbose "PSWindowsUpdate failed: $($_.Exception.Message)"
    return $null
  }
}

# Ensure output path
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

$computer = $env:COMPUTERNAME
$os = (Get-CimInstance Win32_OperatingSystem)
$wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
$wsus = Get-WSUSServer
$times = Get-WULastTimes
$reboot = Get-RebootPending

# pending updates via PSWindowsUpdate (optional) else COM
$pending = @()
if ($UsePSWindowsUpdate) {
  $pending = Get-PendingUpdatesPSWU
  if ($null -eq $pending) { $pending = Get-PendingUpdatesCOM }
} else {
  $pending = Get-PendingUpdatesCOM
}

$pending = $pending | Sort-Object Title
$kbList = ($pending | Select-Object -ExpandProperty KB -ErrorAction Ignore) -join ';'
$sevSummary = (($pending | Where-Object {$_.Severity} | Group-Object Severity | ForEach-Object {"$($_.Name):$($_.Count)"} ) -join ';')

$result = [pscustomobject]@{
  ComputerName    = $computer
  OSVersion       = "$($os.Caption) $($os.Version)"
  LastScanTime    = $times.LastScanTime
  LastInstallTime = $times.LastInstallTime
  PendingCount    = ($pending | Measure-Object).Count
  PendingKBs      = $kbList
  SeveritySummary = $sevSummary
  RebootRequired  = [bool]$reboot
  WUServiceStatus = if ($wuService) { $wuService.Status } else { 'Unknown' }
  WSUS_Server     = $wsus
  Errors          = $null
}

# Write detailed items (per-update) as a sidecar JSON for power users
$stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$base  = Join-Path $OutputPath "PatchAudit-$($computer)-$stamp"

try {
  switch ($Format) {
    'CSV'  { $result | Export-Csv -NoTypeInformation -Path ($base + '.csv') }
    'JSON' { $result | ConvertTo-Json -Depth 4 | Out-File -Encoding UTF8 ($base + '.json') }
    'HTML' {
      $html = $result | ConvertTo-Html -Title "Patch Audit - $computer" -PreContent "<h2>Patch Audit - $computer</h2>"
      $html | Out-File -Encoding UTF8 ($base + '.html')
    }
  }
  # Detailed pending updates dump (JSON)
  $pending | ConvertTo-Json -Depth 5 | Out-File -Encoding UTF8 ($base + '.pending.json')
  Write-Host "Audit complete: $base.*"
} catch {
  $result.Errors = $_.Exception.Message
  $result | ConvertTo-Json -Depth 4 | Out-File -Encoding UTF8 ($base + '.error.json')
  Write-Warning "Audit error: $($result.Errors)"
}