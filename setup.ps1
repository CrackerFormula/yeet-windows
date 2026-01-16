#requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

$logDir = Join-Path $PSScriptRoot 'logs'
New-Item -Path $logDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logDir 'agent.log'
Start-Transcript -Path $logPath -Append | Out-Null

function Assert-CommandExists {
  param([string]$Command)
  if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
    throw "Required command not found: $Command"
  }
}

function Invoke-ExternalCommand {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [string[]]$ArgumentList = @(),
    [string]$Activity = $FilePath
  )
  $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow
  if ($proc.ExitCode -ne 0) {
    throw "$Activity failed with exit code $($proc.ExitCode)."
  }
}

try {
  Assert-CommandExists -Command 'winget'

$apps = @(
  @{ Id = 'Mozilla.Firefox'; Name = 'Firefox' },
  @{ Id = 'Valve.Steam'; Name = 'Steam' },
  @{ Id = 'Discord.Discord'; Name = 'Discord' }
)

foreach ($app in $apps) {
  Write-Host "Installing $($app.Name)..."
  winget install --id $app.Id -e --source winget --accept-source-agreements --accept-package-agreements --silent
  if ($LASTEXITCODE -ne 0) {
    throw "winget install failed for $($app.Id) with exit code $LASTEXITCODE."
  }
}

# Install local packages from installers/manifest.json
$installersDir = Join-Path $PSScriptRoot 'installers'
$manifestPath = Join-Path $installersDir 'manifest.json'
if (Test-Path $manifestPath) {
  $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
  foreach ($entry in $manifest) {
    $filePath = Join-Path $installersDir $entry.file
    if (-not (Test-Path $filePath)) {
      Write-Host "Installer not found: $filePath"
      continue
    }
    $ext = [IO.Path]::GetExtension($filePath).ToLowerInvariant()
    $args = if ($entry.args) { $entry.args } else { '' }
    if ($ext -eq '.msi') {
      Write-Host "Installing MSI $($entry.file)..."
      $msiArgs = "/i `"$filePath`" /qn /norestart $args"
      Invoke-ExternalCommand -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Activity "MSI install for $($entry.file)"
    } elseif ($ext -eq '.exe') {
      Write-Host "Installing EXE $($entry.file)..."
      Invoke-ExternalCommand -FilePath $filePath -ArgumentList $args -Activity "EXE install for $($entry.file)"
    } else {
      Write-Host "Unsupported installer type: $filePath"
    }
  }
} else {
  Write-Host "No manifest found at $manifestPath; skipping local installers."
}

# Remove selected built-in apps (best-effort)
$appsToRemove = @(
  'Microsoft.Windows.Ai.Copilot',
  'Microsoft.Windows.Copilot',
  'MicrosoftTeams',
  'MSTeams',
  'Clipchamp.Clipchamp',
  'Microsoft.MicrosoftSolitaireCollection',
  'Microsoft.YourPhone'
)
foreach ($app in $appsToRemove) {
  Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
  Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $app } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
}

# Remove Teams Machine-Wide Installer (if present)
$teamsRegPaths = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$teamsUninstall = Get-ItemProperty -Path $teamsRegPaths -ErrorAction SilentlyContinue |
  Where-Object { $_.DisplayName -like '*Teams Machine-Wide Installer*' } |
  Select-Object -First 1
if ($teamsUninstall -and $teamsUninstall.UninstallString) {
  Write-Host 'Removing Teams Machine-Wide Installer...'
  $uninstallCmd = $teamsUninstall.UninstallString.Trim('"')
  if ($uninstallCmd -match 'msiexec\.exe') {
    $uninstallArgs = $uninstallCmd -replace '.*msiexec\.exe','' -replace '^\s+',''
    Invoke-ExternalCommand -FilePath 'msiexec.exe' -ArgumentList "$uninstallArgs /qn /norestart" -Activity 'Teams Machine-Wide Uninstall'
  } else {
    Invoke-ExternalCommand -FilePath $uninstallCmd -Activity 'Teams Machine-Wide Uninstall'
  }
}

# Remove OneDrive (best-effort)
Stop-Process -Name 'OneDrive' -Force -ErrorAction SilentlyContinue
$oneDriveSetup = @(
  "$env:SystemRoot\System32\OneDriveSetup.exe",
  "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
)
foreach ($path in $oneDriveSetup) {
  if (Test-Path $path) {
    Start-Process -FilePath $path -ArgumentList '/uninstall' -Wait -WindowStyle Hidden
  }
}
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Force | Out-Null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -Name 'DisableFileSyncNGSC' -Type DWord -Value 1
Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDrive' -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDrive' -ErrorAction SilentlyContinue

# Disable Widgets (Windows 11) and News & Interests (Windows 10 policy)
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Force | Out-Null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Name 'AllowNewsAndInterests' -Type DWord -Value 0

New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Force | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'TaskbarDa' -Type DWord -Value 0

# Disable Cortana
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Force | Out-Null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name 'AllowCortana' -Type DWord -Value 0

# Disable telemetry (policy)
New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Force | Out-Null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name 'AllowTelemetry' -Type DWord -Value 0
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name 'DoNotShowFeedbackNotifications' -Type DWord -Value 1

# Set dark mode for apps and system (current user)
New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Force | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme' -Type DWord -Value 0
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'SystemUsesLightTheme' -Type DWord -Value 0

# Disable feedback prompts
New-Item -Path 'HKCU:\Software\Microsoft\Siuf\Rules' -Force | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Siuf\Rules' -Name 'NumberOfSIUFInPeriod' -Type DWord -Value 0
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Siuf\Rules' -Name 'PeriodInNanoSeconds' -Type QWord -Value 0

# Disable telemetry-related scheduled tasks (best-effort)
$telemetryTasks = @(
  '\Microsoft\Windows\Application Experience\ProgramDataUpdater',
  '\Microsoft\Windows\Application Experience\StartupAppTask',
  '\Microsoft\Windows\Autochk\Proxy',
  '\Microsoft\Windows\Customer Experience Improvement Program\Consolidator',
  '\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask',
  '\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip',
  '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector'
)
foreach ($task in $telemetryTasks) {
  schtasks /Change /TN $task /Disable | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to disable scheduled task: $task (exit code $LASTEXITCODE)."
  }
}

  Write-Host 'Done. You may need to sign out or restart Explorer for taskbar changes to apply.'
} finally {
  try {
    Stop-Transcript | Out-Null
  } catch {
    Write-Host 'Transcript not active; skipping Stop-Transcript.'
  }
}
