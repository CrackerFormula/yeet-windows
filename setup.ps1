$ErrorActionPreference = 'Continue'

# Self-elevate if needed.
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
  $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`"")
  Start-Process -FilePath 'powershell.exe' -ArgumentList $argsList -Verb RunAs
  exit 0
}

$logDir = Join-Path $PSScriptRoot 'logs'
New-Item -Path $logDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logDir 'setup.log'
Start-Transcript -Path $logPath -Append | Out-Null

try {

  # ── Install apps via winget ───────────────────────────────────────────────
  $packagesJson = Join-Path $PSScriptRoot 'packages.json'
  if (Test-Path $packagesJson) {
    Write-Host "`n[1/5] Installing apps from packages.json..." -ForegroundColor Cyan
    winget import -i $packagesJson --accept-source-agreements --accept-package-agreements --ignore-unavailable --ignore-versions
  } else {
    Write-Host "packages.json not found, skipping app install." -ForegroundColor Yellow
  }

  # ── Install local packages from installers/manifest.json ─────────────────
  $installersDir = Join-Path $PSScriptRoot 'installers'
  $manifestPath  = Join-Path $installersDir 'manifest.json'
  if (Test-Path $manifestPath) {
    Write-Host "`n[2/5] Installing local packages..." -ForegroundColor Cyan
    $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
    foreach ($entry in $manifest) {
      $filePath = Join-Path $installersDir $entry.file
      if (-not (Test-Path $filePath)) { Write-Host "  Not found: $filePath"; continue }
      $ext  = [IO.Path]::GetExtension($filePath).ToLowerInvariant()
      $args = if ($entry.args) { $entry.args } else { @() }
      if ($ext -eq '.msi') {
        Write-Host "  Installing MSI: $($entry.file)"
        Start-Process msiexec.exe -ArgumentList "/i `"$filePath`" /qn /norestart $args" -Wait -NoNewWindow
      } elseif ($ext -eq '.exe') {
        Write-Host "  Installing EXE: $($entry.file)"
        Start-Process $filePath -ArgumentList $args -Wait -NoNewWindow
      }
    }
  }

  # ── Remove bloat ──────────────────────────────────────────────────────────
  Write-Host "`n[3/5] Removing bloat..." -ForegroundColor Cyan
  $appsToRemove = @(
    'Microsoft.Windows.Ai.Copilot'
    'Microsoft.Windows.Copilot'
    'MicrosoftTeams'
    'MSTeams'
    'Clipchamp.Clipchamp'
    'Microsoft.MicrosoftSolitaireCollection'
    'Microsoft.YourPhone'
    'Microsoft.BingWeather'
    'Microsoft.BingNews'
    'Microsoft.WindowsMaps'
    'Microsoft.MixedReality.Portal'
    'Microsoft.People'
    'microsoft.windowscommunicationsapps'
    'Microsoft.MicrosoftOfficeHub'
    'Microsoft.Getstarted'
    'Microsoft.WindowsFeedbackHub'
    'Microsoft.ZuneMusic'
    'Microsoft.ZuneVideo'
    'MicrosoftWindows.Client.WebExperience'
  )
  foreach ($app in $appsToRemove) {
    Write-Host "  Removing: $app"
    Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object DisplayName -eq $app | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  }

  # Remove Teams Machine-Wide Installer
  $teamsUninstall = Get-ItemProperty -Path @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  ) -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*Teams Machine-Wide*' } | Select-Object -First 1
  if ($teamsUninstall?.UninstallString) {
    $cmd = $teamsUninstall.UninstallString.Trim('"')
    if ($cmd -match 'msiexec') {
      Start-Process msiexec.exe -ArgumentList (($cmd -replace '.*msiexec\.exe\s*','') + ' /qn /norestart') -Wait -NoNewWindow
    } else { Start-Process $cmd -Wait -NoNewWindow }
  }

  # Remove OneDrive
  Write-Host "  Removing OneDrive..."
  Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
  foreach ($p in @("$env:SystemRoot\System32\OneDriveSetup.exe","$env:SystemRoot\SysWOW64\OneDriveSetup.exe")) {
    if (Test-Path $p) { Start-Process $p -ArgumentList '/uninstall' -Wait -WindowStyle Hidden }
  }
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Name DisableFileSyncNGSC -Type DWord -Value 1
  Remove-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name OneDrive -ErrorAction SilentlyContinue
  Remove-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name OneDrive -ErrorAction SilentlyContinue

  # ── Registry tweaks ───────────────────────────────────────────────────────
  Write-Host "`n[4/5] Applying registry tweaks..." -ForegroundColor Cyan

  # Disable widgets
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Name AllowNewsAndInterests -Type DWord -Value 0
  New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Force | Out-Null
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarDa -Type DWord -Value 0

  # Disable Cortana
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name AllowCortana -Type DWord -Value 0

  # Disable telemetry
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name AllowTelemetry -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name DoNotShowFeedbackNotifications -Type DWord -Value 1

  # Block bloat reinstall
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Name DisableWindowsConsumerFeatures -Type DWord -Value 1

  # Disable Game Bar / DVR (keep Xbox app installed)
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR' -Name AllowGameDVR -Type DWord -Value 0
  New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' -Force | Out-Null
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' -Name AppCaptureEnabled -Type DWord -Value 0

  # Dark mode
  New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Force | Out-Null
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -Type DWord -Value 0
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name SystemUsesLightTheme -Type DWord -Value 0

  # Disable feedback
  New-Item 'HKCU:\Software\Microsoft\Siuf\Rules' -Force | Out-Null
  Set-ItemProperty 'HKCU:\Software\Microsoft\Siuf\Rules' -Name NumberOfSIUFInPeriod -Type DWord -Value 0
  Set-ItemProperty 'HKCU:\Software\Microsoft\Siuf\Rules' -Name PeriodInNanoSeconds -Type QWord -Value 0

  # ── Disable telemetry tasks ───────────────────────────────────────────────
  Write-Host "`n[5/5] Disabling telemetry tasks..." -ForegroundColor Cyan
  @(
    '\Microsoft\Windows\Application Experience\ProgramDataUpdater'
    '\Microsoft\Windows\Application Experience\StartupAppTask'
    '\Microsoft\Windows\Autochk\Proxy'
    '\Microsoft\Windows\Customer Experience Improvement Program\Consolidator'
    '\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask'
    '\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip'
    '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector'
  ) | ForEach-Object {
    schtasks /Change /TN $_ /Disable 2>$null | Out-Null
  }

  Write-Host "`nDone! You may need to restart for all changes to apply." -ForegroundColor Green

} finally {
  try { Stop-Transcript | Out-Null } catch {}
}
