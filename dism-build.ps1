#Requires -RunAsAdministrator
$ErrorActionPreference = 'Continue'

$ISO         = "C:\Users\Bear\Desktop\WindowsMain.iso"
$MountDir    = "C:\WinMount"
$WimDir      = "C:\WimWork"
$WimFile     = "$WimDir\install.wim"
$DriversDir  = "C:\Users\Bear\Desktop\drivers"
$OutputISO   = "C:\Users\Bear\Desktop\WindowsClean.iso"

# ── 1. Mount ISO and extract install.wim ─────────────────────────────────────
Write-Host "`n[1/6] Mounting ISO..." -ForegroundColor Cyan
$mount = Mount-DiskImage -ImagePath $ISO -PassThru
$driveLetter = ($mount | Get-Volume).DriveLetter + ":"
New-Item -Path $MountDir, $WimDir -ItemType Directory -Force | Out-Null

Write-Host "Copying install.wim (may take a minute)..."
$src = "$driveLetter\sources\install.wim"
$esd = "$driveLetter\sources\install.esd"
if (Test-Path $src) {
    Copy-Item $src $WimFile
} elseif (Test-Path $esd) {
    Write-Host "Found install.esd — converting to .wim..."
    dism /Export-Image /SourceImageFile:$esd /SourceIndex:1 /DestinationImageFile:$WimFile /Compress:max /CheckIntegrity
} else {
    throw "No install.wim or install.esd found in ISO"
}

# Show available editions
Write-Host "`nAvailable editions:" -ForegroundColor Yellow
dism /Get-WimInfo /WimFile:$WimFile

$index = Read-Host "`nEnter the index number for the edition you want"

# ── 2. Mount WIM ─────────────────────────────────────────────────────────────
Write-Host "`n[2/6] Mounting WIM index $index..." -ForegroundColor Cyan
attrib -r $WimFile
dism /Mount-Image /ImageFile:$WimFile /Index:$index /MountDir:$MountDir

# ── 3. Integrate drivers ──────────────────────────────────────────────────────
Write-Host "`n[3/6] Integrating drivers from $DriversDir..." -ForegroundColor Cyan
dism /Image:$MountDir /Add-Driver /Driver:$DriversDir /Recurse /ForceUnsigned

# ── 4. Remove bloat ───────────────────────────────────────────────────────────
Write-Host "`n[4/6] Removing bloat..." -ForegroundColor Cyan
$appsToRemove = @(
    'Microsoft.Windows.Ai.Copilot'
    'Microsoft.Copilot'
    'MicrosoftTeams'
    'Microsoft.Teams'
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
    'Microsoft.Xbox.TCUI'
    'Microsoft.XboxApp'
    'Microsoft.XboxGameOverlay'
    'Microsoft.XboxSpeechToTextOverlay'
)

foreach ($app in $appsToRemove) {
    $pkg = dism /Image:$MountDir /Get-ProvisionedAppxPackages | 
           Select-String $app | Select-Object -First 1
    if ($pkg) {
        $pkgName = ($pkg -split ": ")[-1].Trim()
        Write-Host "  Removing $pkgName"
        dism /Image:$MountDir /Remove-ProvisionedAppxPackage /PackageName:$pkgName
    } else {
        Write-Host "  Not found (skipping): $app" -ForegroundColor DarkGray
    }
}

# ── 5. Disable Game Bar via registry ─────────────────────────────────────────
Write-Host "`n[5/6] Disabling Game Bar/DVR..." -ForegroundColor Cyan
$software = "$MountDir\Windows\System32\config\SOFTWARE"
reg load HKLM\OfflineHive $software | Out-Null
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\GameDVR" /v AllowGameDVR /t REG_DWORD /d 0 /f | Out-Null
reg add "HKLM\OfflineHive\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f | Out-Null
reg unload HKLM\OfflineHive | Out-Null

# Block bloat reinstall
reg load HKLM\OfflineHive $software | Out-Null
reg add "HKLM\OfflineHive\Policies\Microsoft\Windows\CloudContent" /v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f | Out-Null
reg unload HKLM\OfflineHive | Out-Null

# ── 6. Commit and rebuild ISO ─────────────────────────────────────────────────
Write-Host "`n[6/6] Committing image..." -ForegroundColor Cyan
dism /Unmount-Image /MountDir:$MountDir /Commit
dism /Export-Image /SourceImageFile:$WimFile /SourceIndex:$index /DestinationImageFile:"$WimDir\install_clean.wim" /Compress:max

Write-Host "`nDone! Clean WIM at: $WimDir\install_clean.wim" -ForegroundColor Green
Write-Host "Use Rufus to create a bootable USB with this WIM." -ForegroundColor Green
Write-Host "Copy setup.ps1 and packages.json to the USB root for post-install." -ForegroundColor Green

Dismount-DiskImage -ImagePath $ISO | Out-Null
