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

  # ---- Install apps via winget ----------------------------------------------------------------------------------------------
  $packagesJson = Join-Path $PSScriptRoot 'packages.json'
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "`n[1/5] winget not found - skipping app installs. Run setup.ps1 again after Windows Update completes." -ForegroundColor Yellow
  } elseif (Test-Path $packagesJson) {
    Write-Host "`n[1/5] Installing apps from packages.json..." -ForegroundColor Cyan
    winget import -i $packagesJson --accept-source-agreements --accept-package-agreements --ignore-unavailable --ignore-versions
  } else {
    Write-Host "`n[1/5] packages.json not found, skipping app install." -ForegroundColor Yellow
  }

  # ---- Install local packages from installers/manifest.json ----------------------------------
  $installersDir = Join-Path $PSScriptRoot 'installers'
  $manifestPath  = Join-Path $installersDir 'manifest.json'
  if (Test-Path $manifestPath) {
    Write-Host "`n[2/5] Installing local packages..." -ForegroundColor Cyan
    $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
    foreach ($entry in $manifest) {
      $filePath = Join-Path $installersDir $entry.file
      if (-not (Test-Path $filePath)) { Write-Host "  Not found: $filePath"; continue }
      $ext         = [IO.Path]::GetExtension($filePath).ToLowerInvariant()
      $installArgs = if ($entry.args) { $entry.args } else { @() }
      if ($ext -eq '.msi') {
        Write-Host "  Installing MSI: $($entry.file)"
        Start-Process msiexec.exe -ArgumentList "/i `"$filePath`" /qn /norestart $installArgs" -Wait -NoNewWindow
      } elseif ($ext -eq '.exe') {
        Write-Host "  Installing EXE: $($entry.file)"
        Start-Process $filePath -ArgumentList $installArgs -Wait -NoNewWindow
      }
    }
  }

  # ---- Remove bloat --------------------------------------------------------------------------------------------------------------------
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
    'Microsoft.SkypeApp'
    'Microsoft.Microsoft3DViewer'
    'Microsoft.MSPaint3D'
    'MicrosoftCorporationII.QuickAssist'
    'Microsoft.PowerAutomateDesktop'
  )
  foreach ($app in $appsToRemove) {
    Write-Host "  Removing: $app"
    Get-AppxPackage -Name "*$app*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$app*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  }

  # Widgets: also remove via winget as fallback
  Write-Host "  Removing Widgets (winget fallback)..."
  winget uninstall --name "Windows Web Experience Pack" --silent --accept-source-agreements 2>$null | Out-Null

  # Remove Teams Machine-Wide Installer
  $teamsUninstall = Get-ItemProperty -Path @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  ) -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*Teams Machine-Wide*' } | Select-Object -First 1
  if ($teamsUninstall -and $teamsUninstall.UninstallString) {
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
  winget uninstall --name "Microsoft OneDrive" --silent --accept-source-agreements 2>$null | Out-Null
  Get-AppxPackage -Name "*OneDrive*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Name DisableFileSyncNGSC -Type DWord -Value 1
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Name PreventNetworkTrafficPreUserSignIn -Type DWord -Value 1
  Remove-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name OneDrive -ErrorAction SilentlyContinue
  Remove-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name OneDrive -ErrorAction SilentlyContinue

  # Remove Copilot
  Write-Host "  Removing Copilot..."
  Get-AppxPackage -Name "*Microsoft.Windows.Ai.Copilot*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxPackage -Name "*Microsoft.Copilot*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Copilot*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  winget uninstall --name "Copilot" --silent --accept-source-agreements 2>$null | Out-Null
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' -Name TurnOffWindowsCopilot -Type DWord -Value 1
  # Hide Copilot taskbar button
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction SilentlyContinue -Name ShowCopilotButton -Type DWord -Value 0

  # Remove Solitaire
  Write-Host "  Removing Solitaire..."
  Get-AppxPackage -Name "*MicrosoftSolitaireCollection*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Solitaire*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  winget uninstall --name "Microsoft Solitaire Collection" --silent --accept-source-agreements 2>$null | Out-Null

  # Remove remaining bloat via winget fallback
  @(
    'Skype'
    'Microsoft 3D Viewer'
    'Paint 3D'
    'Quick Assist'
    'Power Automate'
    'Microsoft Teams'
  ) | ForEach-Object {
    Write-Host "  Removing (winget): $_"
    winget uninstall --name $_ --silent --accept-source-agreements 2>$null | Out-Null
  }

  # Remove Teams Chat taskbar integration
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat' -Name ChatIcon -Type DWord -Value 3

  # ---- Registry tweaks --------------------------------------------------------------------------------------------------------------
  Write-Host "`n[4/5] Applying registry tweaks..." -ForegroundColor Cyan

  # Suppress Edge (keep installed for WebView2, make it invisible)
  Write-Host "  Suppressing Edge..."
  # Disable Edge startup boost and background running
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name StartupBoostEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name BackgroundModeEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name HideFirstRunExperience -Type DWord -Value 1
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name ShowHomeButton -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name EdgeShoppingAssistantEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name PersonalizationReportingEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name EdgeCollectionsEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name ShowMicrosoftRewards -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name SpotlightExperiencesAndRecommendationsEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name NewTabPageContentEnabled -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name NewTabPageQuickLinksEnabled -Type DWord -Value 0
  # Disable Edge autostart on login
  (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue).PSObject.Properties |
    Where-Object { $_.Name -like 'MicrosoftEdgeAutoLaunch*' } |
    ForEach-Object { Remove-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $_.Name -ErrorAction SilentlyContinue }
  (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue).PSObject.Properties |
    Where-Object { $_.Name -like '*Edge*' } |
    ForEach-Object { Remove-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name $_.Name -ErrorAction SilentlyContinue }
  # Unpin Edge from taskbar
  $edgeLnk = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk"
  if (Test-Path $edgeLnk) { Remove-Item $edgeLnk -Force -ErrorAction SilentlyContinue }

  # Set Firefox as default browser + redirect Start menu web searches to Firefox
  Write-Host "  Setting Firefox as default browser..."
  # Force Start menu search to open results in default browser (not Edge)
  if (-not (Test-Path 'HKCU:\Software\Policies\Microsoft\Windows\Explorer')) { New-Item 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Policies\Microsoft\Windows\Explorer' -ErrorAction SilentlyContinue -Name DisableSearchBoxSuggestions -Type DWord -Value 1
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name DisableWebSearch -Type DWord -Value 1
  # Redirect Edge search protocol to Firefox
  if (-not (Test-Path 'HKCU:\Software\Classes\MSEdgeHTM\shell\open\command')) { New-Item 'HKCU:\Software\Classes\MSEdgeHTM\shell\open\command' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Classes\MSEdgeHTM\shell\open\command' -ErrorAction SilentlyContinue -Name '(Default)' -Value '"C:\Program Files\Mozilla Firefox\firefox.exe" -osint -url "%1"'
  # Set Firefox as default via shell (user must confirm in Settings on first run - Windows 11 enforces this)
  Write-Host "  NOTE: Open Settings > Apps > Default Apps > Firefox to finalize browser default." -ForegroundColor Yellow

  # Disable widgets
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Name AllowNewsAndInterests -Type DWord -Value 0
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction SilentlyContinue -Name TaskbarDa -Type DWord -Value 0

  # Remove + disable Cortana
  Write-Host "  Removing Cortana..."
  Get-AppxPackage -Name "*Microsoft.549981C3F5F10*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Cortana*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  winget uninstall --name "Cortana" --silent --accept-source-agreements 2>$null | Out-Null
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name AllowCortana -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name AllowCortanaAboveLock -Type DWord -Value 0

  # Remove feedback hub + disable telemetry
  Write-Host "  Removing Feedback Hub..."
  Get-AppxPackage -Name "*Microsoft.WindowsFeedbackHub*" -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*FeedbackHub*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
  winget uninstall --name "Feedback Hub" --silent --accept-source-agreements 2>$null | Out-Null

  # Disable telemetry
  Write-Host "  Disabling telemetry..."
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name AllowTelemetry -Type DWord -Value 0
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name DoNotShowFeedbackNotifications -Type DWord -Value 1
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name DisableOneSettingsDownloads -Type DWord -Value 1
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name DisableTelemetryOptInChangeNotification -Type DWord -Value 1
  # Kill DiagTrack service
  Stop-Service -Name DiagTrack -Force -ErrorAction SilentlyContinue
  Set-Service -Name DiagTrack -StartupType Disabled -ErrorAction SilentlyContinue
  Stop-Service -Name dmwappushservice -Force -ErrorAction SilentlyContinue
  Set-Service -Name dmwappushservice -StartupType Disabled -ErrorAction SilentlyContinue

  # Block bloat reinstall
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Name DisableWindowsConsumerFeatures -Type DWord -Value 1

  # Disable Game Bar / DVR (keep Xbox app installed)
  New-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR' -Force | Out-Null
  Set-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR' -Name AllowGameDVR -Type DWord -Value 0
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' -ErrorAction SilentlyContinue -Name AppCaptureEnabled -Type DWord -Value 0

  # Dark mode
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue -Name AppsUseLightTheme -Type DWord -Value 0
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue -Name SystemUsesLightTheme -Type DWord -Value 0

  # Disable feedback
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Siuf\Rules')) { New-Item 'HKCU:\Software\Microsoft\Siuf\Rules' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Siuf\Rules' -ErrorAction SilentlyContinue -Name NumberOfSIUFInPeriod -Type DWord -Value 0
  Set-ItemProperty 'HKCU:\Software\Microsoft\Siuf\Rules' -ErrorAction SilentlyContinue -Name PeriodInNanoSeconds -Type QWord -Value 0

  # ---- Disable background app access globally --------------------------------------------------------------
  Write-Host "`n[4b/5] Disabling background apps..." -ForegroundColor Cyan
  # Global kill switch for background apps (user setting)
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications' -ErrorAction SilentlyContinue -Name GlobalUserDisabled -Type DWord -Value 1
  # Also via search policy
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' -ErrorAction SilentlyContinue -Name BackgroundAppGlobalToggle -Type DWord -Value 0
  # Disable specific known background nuisances
  @(
    'Microsoft.YourPhone_8wekyb3d8bbwe'
    'Microsoft.BingWeather_8wekyb3d8bbwe'
    'Microsoft.BingNews_8wekyb3d8bbwe'
    'Microsoft.GetHelp_8wekyb3d8bbwe'
    'Microsoft.Getstarted_8wekyb3d8bbwe'
    'Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe'
    'Microsoft.Windows.Photos_8wekyb3d8bbwe'
    'Microsoft.WindowsMaps_8wekyb3d8bbwe'
    'MicrosoftWindows.Client.WebExperience_cw5n1h2txyewy'
  ) | ForEach-Object {
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications\$_"
    New-Item $path -Force | Out-Null
    Set-ItemProperty $path -Name Disabled -Type DWord -Value 1
    Set-ItemProperty $path -Name DisabledByUser -Type DWord -Value 1
  }

  # Disable startup delay for apps
  if (-not (Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize')) { New-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize' -Force -ErrorAction SilentlyContinue | Out-Null }
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize' -ErrorAction SilentlyContinue -Name StartupDelayInMSec -Type DWord -Value 0

  # ---- Disable telemetry tasks ----------------------------------------------------------------------------------------------
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
