#Requires -RunAsAdministrator
$ErrorActionPreference = 'Continue'

# ---- WARNING: autounattend.xml creates a local admin account with an EMPTY password.
# ---- Change the password immediately after first boot, or modify the XML below.

param(
    [string]$ISO        = "",
    [string]$DriversDir = "",
    [string]$MountDir   = "C:\WinMount",
    [string]$WimDir     = "C:\WimWork"
)
if (-not $ISO)        { $ISO        = "$env:USERPROFILE\Desktop\WindowsMain.iso" }
if (-not $DriversDir) { $DriversDir = "$env:USERPROFILE\Desktop\drivers" }

$WimFile    = "$WimDir\install.wim"
$OutputWim  = "$WimDir\install_clean.wim"
$scriptSrc  = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- Load shared bloat list --------------------------------------------------
$bloatJsonPath = Join-Path $scriptSrc 'bloat.json'
if (Test-Path $bloatJsonPath) {
  $appsToRemove = Get-Content -Path $bloatJsonPath -Raw | ConvertFrom-Json
} else {
  Write-Host "bloat.json not found, using built-in list." -ForegroundColor Yellow
  $appsToRemove = @(
    'Microsoft.Windows.Ai.Copilot','Microsoft.Windows.Copilot','MicrosoftTeams','MSTeams',
    'Clipchamp.Clipchamp','Microsoft.MicrosoftSolitaireCollection','Microsoft.YourPhone',
    'Microsoft.BingWeather','Microsoft.BingNews','Microsoft.WindowsMaps',
    'Microsoft.MixedReality.Portal','Microsoft.People','microsoft.windowscommunicationsapps',
    'Microsoft.MicrosoftOfficeHub','Microsoft.Getstarted','Microsoft.WindowsFeedbackHub',
    'Microsoft.ZuneMusic','Microsoft.ZuneVideo','MicrosoftWindows.Client.WebExperience',
    'Microsoft.SkypeApp','Microsoft.Microsoft3DViewer','Microsoft.MSPaint3D',
    'MicrosoftCorporationII.QuickAssist','Microsoft.PowerAutomateDesktop'
  )
}

try {

# -- 0. Cleanup from previous run -----------------------------------------------
Write-Host "[0/8] Cleaning up previous run..." -ForegroundColor Cyan
reg unload HKLM\OfflineHive 2>$null | Out-Null
dism /Unmount-Image /MountDir:$MountDir /Discard 2>$null | Out-Null
Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
if (Test-Path $MountDir) { Remove-Item $MountDir -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $WimDir)   { Remove-Item $WimDir   -Recurse -Force -ErrorAction SilentlyContinue }

# -- 1. Mount ISO and extract install file ---------------------------------------
Write-Host "[1/8] Mounting ISO..." -ForegroundColor Cyan
$mount = Mount-DiskImage -ImagePath $ISO -PassThru
$vol = $mount | Get-Volume
if (-not $vol -or -not $vol.DriveLetter) {
    throw "Failed to mount ISO or no drive letter assigned."
}
$drive = $vol.DriveLetter + ":"
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

# Dismount ISO now (we have the WIM extracted); will remount later for ISO build.
Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null

# -- 2. Mount WIM ---------------------------------------------------------------
Write-Host ""
Write-Host "[2/8] Mounting WIM..." -ForegroundColor Cyan
attrib -r $WimFile
dism /Mount-Image /ImageFile:$WimFile /Index:$index /MountDir:$MountDir

# -- 3. Integrate drivers -------------------------------------------------------
Write-Host ""
Write-Host "[3/8] Integrating drivers..." -ForegroundColor Cyan
if (Test-Path $DriversDir) {
    dism /Image:$MountDir /Add-Driver /Driver:$DriversDir /Recurse /ForceUnsigned
    if ($LASTEXITCODE -ne 0) { Write-Host "  WARNING: Driver integration failed (exit $LASTEXITCODE)" -ForegroundColor Yellow }
} else {
    Write-Host "  No drivers directory at $DriversDir -- skipping." -ForegroundColor Yellow
}

# -- 4. Remove bloat -------------------------------------------------------------
Write-Host ""
Write-Host "[4/8] Removing bloat..." -ForegroundColor Cyan

$allPackages = dism /Image:$MountDir /Get-ProvisionedAppxPackages

foreach ($app in $appsToRemove) {
    $pkgLine = $allPackages | Select-String "PackageName" | Where-Object { $_ -match [regex]::Escape($app) } | Select-Object -First 1
    if ($pkgLine) {
        $pkgName = ($pkgLine -split ": ", 2)[-1].Trim()
        Write-Host "  Removing $pkgName"
        dism /Image:$MountDir /Remove-ProvisionedAppxPackage /PackageName:$pkgName
    } else {
        Write-Host "  Not found (skipping): $app" -ForegroundColor DarkGray
    }
}

# -- 5. Bake autounattend.xml into image ----------------------------------------
Write-Host ""
Write-Host "[5/8] Baking autounattend.xml into image..." -ForegroundColor Cyan
$unattendContent = @'
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="windowsPE">
    <component name="Microsoft-Windows-International-Core-WinPE" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <SetupUILanguage>
        <UILanguage>en-US</UILanguage>
      </SetupUILanguage>
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UserLocale>en-US</UserLocale>
    </component>
    <component name="Microsoft-Windows-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <UserData>
        <AcceptEula>true</AcceptEula>
      </UserData>
    </component>
  </settings>
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <OOBE>
        <HideEULAPage>true</HideEULAPage>
        <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <NetworkLocation>Home</NetworkLocation>
        <ProtectYourPC>3</ProtectYourPC>
        <SkipMachineOOBE>true</SkipMachineOOBE>
        <SkipUserOOBE>true</SkipUserOOBE>
      </OOBE>
      <UserAccounts>
        <LocalAccounts>
          <LocalAccount wcm:action="add">
            <Password>
              <Value></Value>
              <PlainText>true</PlainText>
            </Password>
            <DisplayName>Bear</DisplayName>
            <Group>Administrators</Group>
            <Name>Bear</Name>
          </LocalAccount>
        </LocalAccounts>
      </UserAccounts>
      <AutoLogon>
        <Password>
          <Value></Value>
          <PlainText>true</PlainText>
        </Password>
        <Enabled>true</Enabled>
        <LogonCount>1</LogonCount>
        <Username>Bear</Username>
      </AutoLogon>
      <TimeZone>Eastern Standard Time</TimeZone>
    </component>
    <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UserLocale>en-US</UserLocale>
    </component>
  </settings>
</unattend>
'@
$pantherDir = "$MountDir\Windows\Panther"
New-Item -Path $pantherDir -ItemType Directory -Force | Out-Null
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText("$pantherDir\unattend.xml", $unattendContent, $utf8NoBom)
# Also save a copy next to the script for placing on USB root
[System.IO.File]::WriteAllText("$scriptSrc\autounattend.xml", $unattendContent, $utf8NoBom)
Write-Host "  autounattend.xml baked in. Also saved to $scriptSrc\autounattend.xml for Ventoy USB root."

# -- 6. Bake setup.ps1 into image -----------------------------------------------
Write-Host ""
Write-Host "[6/8] Baking setup.ps1 into image..." -ForegroundColor Cyan
$scriptsDir = "$MountDir\Windows\Setup\Scripts"
New-Item -Path $scriptsDir -ItemType Directory -Force | Out-Null
Copy-Item "$scriptSrc\setup.ps1"     "$scriptsDir\setup.ps1"    -Force
Copy-Item "$scriptSrc\packages.json" "$scriptsDir\packages.json" -Force
Copy-Item "$scriptSrc\bloat.json"    "$scriptsDir\bloat.json"    -Force
# Copy installers directory if it exists and has real files
$installersDir = Join-Path $scriptSrc 'installers'
if (Test-Path $installersDir) {
    $realInstallers = Get-ChildItem $installersDir -Include '*.msi','*.exe' -Recurse -ErrorAction SilentlyContinue
    if ($realInstallers) {
        $destInstallers = "$scriptsDir\installers"
        New-Item -Path $destInstallers -ItemType Directory -Force | Out-Null
        Copy-Item "$installersDir\*" $destInstallers -Recurse -Force
        Write-Host "  Copied installers directory into image."
    }
}
# SetupComplete.cmd runs automatically on first boot
@"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%WINDIR%\Setup\Scripts\setup.ps1"
"@ | Set-Content "$scriptsDir\SetupComplete.cmd" -Encoding ASCII

# -- 7. Registry tweaks via offline hive -----------------------------------------
Write-Host ""
Write-Host "[7/8] Applying registry tweaks..." -ForegroundColor Cyan
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

# Unload hive with retry (handles locked handles from AV/indexer)
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
$retries = 0
do {
    reg unload HKLM\OfflineHive 2>$null
    if ($LASTEXITCODE -eq 0) { break }
    $retries++
    Write-Host "  Hive busy, retrying ($retries/5)..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    [GC]::Collect()
} while ($retries -lt 5)
if ($LASTEXITCODE -ne 0) {
    throw "Failed to unload registry hive -- WIM commit would corrupt. Aborting."
}

# -- 8. Commit and export clean WIM + build ISO ---------------------------------
Write-Host ""
Write-Host "[8/8] Committing and exporting clean WIM..." -ForegroundColor Cyan
dism /Unmount-Image /MountDir:$MountDir /Commit
dism /Export-Image /SourceImageFile:$WimFile /SourceIndex:$index /DestinationImageFile:$OutputWim /Compress:max

# Build bootable ISO with scripts baked in
Write-Host ""
Write-Host "Building ISO..." -ForegroundColor Cyan

$IsoStaging = "C:\WimWork\iso-staging"
$OutputISO  = "$env:USERPROFILE\Desktop\WindowsClean.iso"
$oscdimg    = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"

if (-not (Test-Path $oscdimg)) {
    Write-Host "oscdimg not found -- Windows ADK not installed." -ForegroundColor Yellow
    Write-Host "Install ADK from: https://go.microsoft.com/fwlink/?linkid=2243390" -ForegroundColor Yellow
    Write-Host "Then rerun this script. Clean WIM is ready at: $OutputWim" -ForegroundColor Green
} else {
    # Remount ISO to copy boot files
    Write-Host "  Copying ISO boot files to staging..."
    $isoMount = Mount-DiskImage -ImagePath $ISO -PassThru
    $isoVol = $isoMount | Get-Volume
    if (-not $isoVol -or -not $isoVol.DriveLetter) {
        throw "Failed to remount ISO for staging."
    }
    $isoDrive = $isoVol.DriveLetter + ":"
    New-Item -Path $IsoStaging -ItemType Directory -Force | Out-Null
    Copy-Item "$isoDrive\*" $IsoStaging -Recurse -Force

    # Replace install file with our clean WIM
    Remove-Item "$IsoStaging\sources\install.esd" -ErrorAction SilentlyContinue
    Remove-Item "$IsoStaging\sources\install.wim" -ErrorAction SilentlyContinue
    Copy-Item $OutputWim "$IsoStaging\sources\install.wim"

    # Bake in setup.ps1, packages.json, and bloat.json
    Copy-Item "$scriptSrc\setup.ps1"        "$IsoStaging\setup.ps1"
    Copy-Item "$scriptSrc\packages.json"    "$IsoStaging\packages.json"
    Copy-Item "$scriptSrc\bloat.json"       "$IsoStaging\bloat.json"
    Copy-Item "$scriptSrc\autounattend.xml" "$IsoStaging\autounattend.xml"

    Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null

    # Build ISO
    Write-Host "  Building ISO (this may take a few minutes)..."
    $efiBoot  = "$IsoStaging\efi\microsoft\boot\efisys.bin"
    $etfsBoot = "$IsoStaging\boot\etfsboot.com"
    & $oscdimg -m -o -u2 -udfver102 `
        -bootdata:"2#p0,e,b$etfsBoot#pEF,e,b$efiBoot" `
        $IsoStaging $OutputISO

    Write-Host ""
    Write-Host "Done! ISO ready: $OutputISO" -ForegroundColor Green
    Write-Host "Drop it on Ventoy and boot -- setup.ps1 and packages.json are in the root." -ForegroundColor Green
}

} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    throw
} finally {
    # Guaranteed cleanup: unload hive, unmount WIM, dismount ISO
    reg unload HKLM\OfflineHive 2>$null | Out-Null
    dism /Unmount-Image /MountDir:$MountDir /Discard 2>$null | Out-Null
    Dismount-DiskImage -ImagePath $ISO -ErrorAction SilentlyContinue | Out-Null
}
