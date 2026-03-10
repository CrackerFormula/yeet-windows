#Requires -RunAsAdministrator
$ErrorActionPreference = 'Continue'

$ISO        = "C:\Users\Bear\Desktop\WindowsMain.iso"
$MountDir   = "C:\WinMount"
$WimDir     = "C:\WimWork"
$WimFile    = "$WimDir\install.wim"
$DriversDir = "C:\Users\Bear\Desktop\drivers"
$OutputWim  = "$WimDir\install_clean.wim"

# -- 1. Mount ISO and extract install file ------------------------------------
Write-Host "[1/6] Mounting ISO..." -ForegroundColor Cyan
$mount = Mount-DiskImage -ImagePath $ISO -PassThru
$drive = ($mount | Get-Volume).DriveLetter + ":"
New-Item -Path $MountDir, $WimDir -ItemType Directory -Force | Out-Null

$src = "$drive\sources\install.wim"
$esd = "$drive\sources\install.esd"

if (Test-Path $src) {
    Write-Host "Found install.wim -- copying..."
    Copy-Item $src $WimFile
    Write-Host ""
    Write-Host "Available editions:" -ForegroundColor Yellow
    dism /Get-WimInfo /WimFile:$WimFile
    $index = Read-Host "Enter index number for edition you want"
} elseif (Test-Path $esd) {
    Write-Host "Found install.esd -- showing editions..."
    dism /Get-WimInfo /WimFile:$esd
    $index = Read-Host "Enter index number for edition you want"
    Write-Host "Exporting index $index from ESD (takes several minutes)..."
    dism /Export-Image /SourceImageFile:$esd /SourceIndex:$index /DestinationImageFile:$WimFile /Compress:max /CheckIntegrity
    $index = 1
} else {
    throw "No install.wim or install.esd found in ISO at $drive\sources\"
}

# -- 2. Mount WIM -------------------------------------------------------------
Write-Host ""
Write-Host "[2/6] Mounting WIM..." -ForegroundColor Cyan
attrib -r $WimFile
dism /Mount-Image /ImageFile:$WimFile /Index:$index /MountDir:$MountDir

# -- 3. Integrate drivers -----------------------------------------------------
Write-Host ""
Write-Host "[3/6] Integrating drivers from $DriversDir..." -ForegroundColor Cyan
dism /Image:$MountDir /Add-Driver /Driver:$DriversDir /Recurse /ForceUnsigned

# -- 4. Remove bloat ----------------------------------------------------------
Write-Host ""
Write-Host "[4/6] Removing bloat..." -ForegroundColor Cyan
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
    $pkg = dism /Image:$MountDir /Get-ProvisionedAppxPackages | Select-String $app | Select-Object -First 1
    if ($pkg) {
        $pkgName = ($pkg -split ": ")[-1].Trim()
        Write-Host "  Removing $pkgName"
        dism /Image:$MountDir /Remove-ProvisionedAppxPackage /PackageName:$pkgName
    } else {
        Write-Host "  Not found (skipping): $app" -ForegroundColor DarkGray
    }
}

# -- 5. Registry tweaks via offline hive --------------------------------------
Write-Host ""
Write-Host "[5/6] Applying registry tweaks..." -ForegroundColor Cyan
$software = "$MountDir\Windows\System32\config\SOFTWARE"

reg load HKLM\OfflineHive $software | Out-Null

# Disable Game Bar / DVR
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\GameDVR" /v AllowGameDVR /t REG_DWORD /d 0 /f | Out-Null
reg add "HKLM\OfflineHive\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f | Out-Null

# Block bloat reinstall
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\CloudContent" /v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f | Out-Null

# Disable telemetry
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f | Out-Null
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\DataCollection" /v DoNotShowFeedbackNotifications /t REG_DWORD /d 1 /f | Out-Null

# Disable Cortana
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\Windows Search" /v AllowCortana /t REG_DWORD /d 0 /f | Out-Null

# Suppress Edge
reg add "HKLM\OfflineHive\Policies\Microsoft\Edge" /v StartupBoostEnabled /t REG_DWORD /d 0 /f | Out-Null
reg add "HKLM\OfflineHive\Policies\Microsoft\Edge" /v BackgroundModeEnabled /t REG_DWORD /d 0 /f | Out-Null
reg add "HKLM\OfflineHive\Policies\Microsoft\Edge" /v HideFirstRunExperience /t REG_DWORD /d 1 /f | Out-Null

# Disable OneDrive
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\OneDrive" /v DisableFileSyncNGSC /t REG_DWORD /d 1 /f | Out-Null

# Disable Widgets
reg add "HKLM\OfflineHive\Policies\Microsoft\Dsh" /v AllowNewsAndInterests /t REG_DWORD /d 0 /f | Out-Null

# Disable Copilot
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f | Out-Null

reg unload HKLM\OfflineHive | Out-Null

# -- 6. Commit and export clean WIM -------------------------------------------
Write-Host ""
Write-Host "[6/6] Committing and exporting clean WIM..." -ForegroundColor Cyan
dism /Unmount-Image /MountDir:$MountDir /Commit
dism /Export-Image /SourceImageFile:$WimFile /SourceIndex:1 /DestinationImageFile:$OutputWim /Compress:max

Write-Host ""
Write-Host "Done! Clean WIM: $OutputWim" -ForegroundColor Green
Write-Host "Use Rufus to create bootable USB with this WIM." -ForegroundColor Green
Write-Host "Copy setup.ps1 and packages.json to the USB root for post-install." -ForegroundColor Green

Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
