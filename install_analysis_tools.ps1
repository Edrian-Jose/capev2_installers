<#
.SYNOPSIS
    Install analysis tools system-wide for CAPE malware sandbox

.DESCRIPTION
    This script installs common applications that malware targets, configured
    for system-wide installation to survive sysprep. All security features,
    update checks, and telemetry are disabled.

.NOTES
    - Run as Administrator
    - Internet connection required
    - Designed for Windows 10/11
    - Installs to system directories (survives sysprep)
    - Uses Chocolatey package manager

.EXAMPLE
    .\install_analysis_tools.ps1

.EXAMPLE
    .\install_analysis_tools.ps1 -SkipOffice -SkipAdobeReader
#>

[CmdletBinding()]
param(
    [switch]$SkipOffice,
    [switch]$SkipAdobeReader,
    [switch]$SkipBrowsers,
    [switch]$SkipDevelopmentTools,
    [switch]$SkipMediaPlayers,
    [switch]$SkipArchivers
)

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

# Set execution policy
Set-ExecutionPolicy Bypass -Scope Process -Force

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "CAPE Malware Analysis Tools Installer" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

#region Helper Functions

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "==== $Title ====" -ForegroundColor Yellow
}

function Test-CommandExists {
    param([string]$Command)
    return [bool](Get-Command -Name $Command -ErrorAction SilentlyContinue)
}

function Install-ChocoPackage {
    param(
        [string]$PackageName,
        [string]$DisplayName,
        [string[]]$AdditionalArgs = @()
    )

    Write-Host "Installing $DisplayName..." -ForegroundColor Green

    $args = @("install", $PackageName, "-y", "--no-progress", "--ignore-checksums") + $AdditionalArgs

    try {
        & choco @args | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  [+] $DisplayName installed successfully" -ForegroundColor Green
            return $true
        } else {
            Write-Host "  [-] Failed to install $DisplayName" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "  [-] Error installing $DisplayName : $_" -ForegroundColor Red
        return $false
    }
}

function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [object]$Value,
        [string]$Type = "DWord"
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        return $true
    } catch {
        Write-Host "  Warning: Could not set registry value $Path\$Name" -ForegroundColor Yellow
        return $false
    }
}

#endregion

#region Install Chocolatey

Write-Section "Installing Chocolatey Package Manager"

if (Test-CommandExists "choco") {
    Write-Host "Chocolatey already installed" -ForegroundColor Green
} else {
    Write-Host "Installing Chocolatey..." -ForegroundColor Green
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

    if (Test-CommandExists "choco") {
        Write-Host "  [+] Chocolatey installed successfully" -ForegroundColor Green
    } else {
        Write-Error "Failed to install Chocolatey. Exiting."
        Exit 1
    }
}

# Refresh environment variables
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

#endregion

#region Install Python (32-bit for agent compatibility)

Write-Section "Installing Python 3.8 (32-bit)"

Install-ChocoPackage -PackageName "python38" -DisplayName "Python 3.8 32-bit" -AdditionalArgs @("--forcex86", "--params", "/InstallDir:C:\Python38-32")

# Install Python packages for agent
if (Test-Path "C:\Python38-32\python.exe") {
    Write-Host "Installing Python packages..." -ForegroundColor Green
    & C:\Python38-32\python.exe -m pip install --upgrade pip --quiet
    & C:\Python38-32\python.exe -m pip install Pillow pywin32 --quiet
    Write-Host "  [+] Python packages installed" -ForegroundColor Green
}

#endregion

#region Install .NET Frameworks

Write-Section "Installing .NET Frameworks"

Install-ChocoPackage -PackageName "dotnet-sdk" -DisplayName ".NET SDK (Latest)"
Install-ChocoPackage -PackageName "dotnetfx" -DisplayName ".NET Framework 4.8"
Install-ChocoPackage -PackageName "dotnet-4.5.2-devpack" -DisplayName ".NET Framework 4.5.2"
Install-ChocoPackage -PackageName "dotnet-3.5" -DisplayName ".NET Framework 3.5"

#endregion

#region Install Java

Write-Section "Installing Java Runtime Environments"

Install-ChocoPackage -PackageName "openjdk" -DisplayName "OpenJDK (Latest)"
Install-ChocoPackage -PackageName "openjdk11" -DisplayName "OpenJDK 11"
Install-ChocoPackage -PackageName "openjdk8" -DisplayName "OpenJDK 8"

# Disable Java auto-update
Write-Host "Disabling Java auto-update..." -ForegroundColor Green
Set-RegistryValue -Path "HKLM:\SOFTWARE\JavaSoft\Java Update\Policy" -Name "EnableJavaUpdate" -Value 0
Set-RegistryValue -Path "HKLM:\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy" -Name "EnableJavaUpdate" -Value 0

#endregion

#region Install Development Tools

if (-not $SkipDevelopmentTools) {
    Write-Section "Installing Development Tools"

    Install-ChocoPackage -PackageName "git" -DisplayName "Git" -AdditionalArgs @("--params", "/GitAndUnixToolsOnPath /NoAutoCrlf")
    Install-ChocoPackage -PackageName "nodejs-lts" -DisplayName "Node.js LTS"
    Install-ChocoPackage -PackageName "yarn" -DisplayName "Yarn"
    Install-ChocoPackage -PackageName "ruby" -DisplayName "Ruby"
    Install-ChocoPackage -PackageName "go" -DisplayName "Go (Golang)"
    Install-ChocoPackage -PackageName "php" -DisplayName "PHP"
    Install-ChocoPackage -PackageName "composer" -DisplayName "Composer (PHP)"

    # Disable Git auto-update
    if (Test-Path "C:\Program Files\Git\cmd\git.exe") {
        & git config --system core.autocrlf false
        & git config --system credential.helper wincred
    }

    # Disable npm update check
    if (Test-CommandExists "npm") {
        & npm config set update-notifier false --global
    }
}

#endregion

#region Install Web Browsers

if (-not $SkipBrowsers) {
    Write-Section "Installing Web Browsers"

    Install-ChocoPackage -PackageName "googlechrome" -DisplayName "Google Chrome" -AdditionalArgs @("--ignore-checksums")
    Install-ChocoPackage -PackageName "firefox" -DisplayName "Mozilla Firefox"

    # Configure Chrome (disable updates, first-run)
    Write-Host "Configuring Chrome..." -ForegroundColor Green

    # Disable Chrome auto-update
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Update" -Name "AutoUpdateCheckPeriodMinutes" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Update" -Name "UpdateDefault" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "DefaultBrowserSettingEnabled" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "BackgroundModeEnabled" -Value 0

    # Disable Chrome sync, sign-in, first-run
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "SyncDisabled" -Value 1
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "BrowserSignin" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "PromotionalTabsEnabled" -Value 0

    # Master preferences for Chrome
    $chromePath = "C:\Program Files\Google\Chrome\Application"
    if (Test-Path $chromePath) {
        $masterPrefs = @'
{
  "homepage": "about:blank",
  "homepage_is_newtabpage": true,
  "browser": {
    "show_home_button": false,
    "check_default_browser": false
  },
  "first_run_tabs": [],
  "sync_promo": {
    "show_on_first_run_allowed": false
  },
  "distribution": {
    "skip_first_run_ui": true,
    "show_welcome_page": false,
    "import_bookmarks": false,
    "import_history": false,
    "import_search_engine": false,
    "suppress_first_run_bubble": true,
    "do_not_create_desktop_shortcut": true,
    "do_not_create_quick_launch_shortcut": true,
    "do_not_create_taskbar_shortcut": true,
    "do_not_launch_chrome": true,
    "do_not_register_for_update_launch": true,
    "make_chrome_default": false
  }
}
'@
        $masterPrefs | Out-File "$chromePath\master_preferences" -Encoding ASCII -Force
        Write-Host "  [+] Chrome configured" -ForegroundColor Green
    }

    # Configure Firefox (disable updates, first-run)
    Write-Host "Configuring Firefox..." -ForegroundColor Green

    $firefoxPath = "C:\Program Files\Mozilla Firefox"
    if (Test-Path $firefoxPath) {
        # Create policies.json
        $policiesDir = "$firefoxPath\distribution"
        New-Item -ItemType Directory -Path $policiesDir -Force | Out-Null

        $policies = @'
{
  "policies": {
    "DisableAppUpdate": true,
    "DisableSystemAddonUpdate": true,
    "DisableTelemetry": true,
    "DisableFirefoxStudies": true,
    "DisablePocket": true,
    "DontCheckDefaultBrowser": true,
    "OverrideFirstRunPage": "",
    "OverridePostUpdatePage": "",
    "NoDefaultBookmarks": true,
    "PasswordManagerEnabled": false,
    "DisableProfileImport": true,
    "DisableFirefoxAccounts": true,
    "FirefoxHome": {
      "Search": false,
      "TopSites": false,
      "Highlights": false,
      "Pocket": false,
      "Snippets": false
    }
  }
}
'@
        $policies | Out-File "$policiesDir\policies.json" -Encoding ASCII -Force
        Write-Host "  [+] Firefox configured" -ForegroundColor Green
    }
}

#endregion

#region Install Adobe Reader

if (-not $SkipAdobeReader) {
    Write-Section "Installing Adobe Reader DC"

    Install-ChocoPackage -PackageName "adobereader" -DisplayName "Adobe Reader DC"

    # Configure Adobe Reader (disable updates, protected mode, cloud)
    Write-Host "Configuring Adobe Reader..." -ForegroundColor Green

    # Disable updates
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bUpdater" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\WOW6432Node\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bUpdater" -Value 0

    # Disable Protected Mode (allows malware to run)
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bProtectedMode" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "iProtectedView" -Value 0

    # Disable cloud features
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown\cCloud" -Name "bAdobeSendPluginToggle" -Value 0

    # Disable upsell
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bAcroSuppressUpsell" -Value 1

    # Disable usage tracking
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bUsageMeasurement" -Value 0

    # Enable JavaScript (malware often uses it)
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown" -Name "bDisableJavaScript" -Value 0

    Write-Host "  [+] Adobe Reader configured" -ForegroundColor Green
}

#endregion

#region Install Office (if available)

if (-not $SkipOffice) {
    Write-Section "Installing Microsoft Office"

    Write-Host "NOTE: Office requires manual installation or volume license" -ForegroundColor Yellow
    Write-Host "If you have Office 2016/2019/365, install it manually, then run:" -ForegroundColor Yellow
    Write-Host "  .\configure_office.ps1" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Installing Office alternatives..." -ForegroundColor Green

    # Install LibreOffice as alternative
    Install-ChocoPackage -PackageName "libreoffice-fresh" -DisplayName "LibreOffice"
}

#endregion

#region Install Archivers

if (-not $SkipArchivers) {
    Write-Section "Installing Archive Tools"

    Install-ChocoPackage -PackageName "7zip" -DisplayName "7-Zip"
    Install-ChocoPackage -PackageName "winrar" -DisplayName "WinRAR"

    # Set 7-Zip as default for common archive formats
    if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
        $extensions = @(".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".iso")
        foreach ($ext in $extensions) {
            cmd /c "assoc $ext=7-Zip.$($ext.TrimStart('.'))" 2>&1 | Out-Null
        }
        Write-Host "  [+] 7-Zip set as default archiver" -ForegroundColor Green
    }
}

#endregion

#region Install Media Players

if (-not $SkipMediaPlayers) {
    Write-Section "Installing Media Players"

    Install-ChocoPackage -PackageName "vlc" -DisplayName "VLC Media Player"
    Install-ChocoPackage -PackageName "k-litecodecpackfull" -DisplayName "K-Lite Codec Pack"

    # Disable VLC updates
    if (Test-Path "C:\Program Files\VideoLAN\VLC") {
        Set-RegistryValue -Path "HKLM:\SOFTWARE\VideoLAN\VLC" -Name "CheckUpdates" -Value 0
    }
}

#endregion

#region Install Additional Common Tools

Write-Section "Installing Additional Common Tools"

Install-ChocoPackage -PackageName "notepadplusplus" -DisplayName "Notepad++"
Install-ChocoPackage -PackageName "putty" -DisplayName "PuTTY"
Install-ChocoPackage -PackageName "winscp" -DisplayName "WinSCP"
Install-ChocoPackage -PackageName "filezilla" -DisplayName "FileZilla"
Install-ChocoPackage -PackageName "postman" -DisplayName "Postman"
Install-ChocoPackage -PackageName "curl" -DisplayName "cURL"
Install-ChocoPackage -PackageName "wget" -DisplayName "Wget"

#endregion

#region Install Flash Player (for old malware)

Write-Section "Installing Adobe Flash Player"

Write-Host "Installing Flash Player (for legacy malware analysis)..." -ForegroundColor Green
Install-ChocoPackage -PackageName "flashplayerplugin" -DisplayName "Flash Player Plugin" -AdditionalArgs @("--ignore-checksums")

#endregion

#region Configure Windows System Settings

Write-Section "Configuring Windows System Settings"

Write-Host "Disabling Windows security features..." -ForegroundColor Green

# Disable Windows Defender
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender" -Name "DisableAntiSpyware" -Value 1
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" -Name "DisableRealtimeMonitoring" -Value 1
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" -Name "DisableBehaviorMonitoring" -Value 1
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" -Name "DisableIOAVProtection" -Value 1
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" -Name "DisableScriptScanning" -Value 1

# Disable Windows Defender via PowerShell (immediate effect)
try {
    Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction SilentlyContinue
    Set-MpPreference -DisableBehaviorMonitoring $true -ErrorAction SilentlyContinue
    Set-MpPreference -DisableIOAVProtection $true -ErrorAction SilentlyContinue
    Set-MpPreference -DisableScriptScanning $true -ErrorAction SilentlyContinue
    Set-MpPreference -DisableBlockAtFirstSeen $true -ErrorAction SilentlyContinue
} catch {
    Write-Host "  Note: Some Defender settings may require reboot" -ForegroundColor Yellow
}

# Disable Windows Firewall
Write-Host "Disabling Windows Firewall..." -ForegroundColor Green
Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False

# Disable UAC
Write-Host "Disabling UAC..." -ForegroundColor Green
Set-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 0

# Disable Windows Updates
Write-Host "Disabling Windows Update..." -ForegroundColor Green
Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
Set-Service wuauserv -StartupType Disabled -ErrorAction SilentlyContinue

# Disable Windows Error Reporting
Set-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" -Name "Disabled" -Value 1

# Disable SmartScreen
Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableSmartScreen" -Value 0
Set-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Value "Off" -Type "String"

# Show file extensions
Set-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0

# Disable crash auto-restart
Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -Name "AutoReboot" -Value 0

# Disable hibernation (saves disk space)
Set-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HibernateEnabled" -Value 0

# Disable system restore (saves disk space)
Disable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue

Write-Host "  [+] Windows security features disabled" -ForegroundColor Green

#endregion

#region Configure Office (if installed)

Write-Section "Configuring Microsoft Office (if installed)"

$officePaths = @(
    "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0",
    "HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0",
    "HKLM:\SOFTWARE\Policies\Microsoft\Office\14.0"
)

foreach ($officePath in $officePaths) {
    if (Test-Path $officePath) {
        Write-Host "Configuring Office..." -ForegroundColor Green

        # Disable macro warnings (DANGEROUS - only for isolated analysis!)
        Set-RegistryValue -Path "$officePath\Word\Security" -Name "VBAWarnings" -Value 1
        Set-RegistryValue -Path "$officePath\Excel\Security" -Name "VBAWarnings" -Value 1
        Set-RegistryValue -Path "$officePath\PowerPoint\Security" -Name "VBAWarnings" -Value 1

        # Disable Protected View
        Set-RegistryValue -Path "$officePath\Word\Security\ProtectedView" -Name "DisableAttachmentsInPV" -Value 1
        Set-RegistryValue -Path "$officePath\Word\Security\ProtectedView" -Name "DisableInternetFilesInPV" -Value 1
        Set-RegistryValue -Path "$officePath\Word\Security\ProtectedView" -Name "DisableUnsafeLocationsInPV" -Value 1
        Set-RegistryValue -Path "$officePath\Excel\Security\ProtectedView" -Name "DisableAttachmentsInPV" -Value 1
        Set-RegistryValue -Path "$officePath\Excel\Security\ProtectedView" -Name "DisableInternetFilesInPV" -Value 1
        Set-RegistryValue -Path "$officePath\Excel\Security\ProtectedView" -Name "DisableUnsafeLocationsInPV" -Value 1
        Set-RegistryValue -Path "$officePath\PowerPoint\Security\ProtectedView" -Name "DisableAttachmentsInPV" -Value 1
        Set-RegistryValue -Path "$officePath\PowerPoint\Security\ProtectedView" -Name "DisableInternetFilesInPV" -Value 1
        Set-RegistryValue -Path "$officePath\PowerPoint\Security\ProtectedView" -Name "DisableUnsafeLocationsInPV" -Value 1

        # Disable Office telemetry
        Set-RegistryValue -Path "$officePath\Common" -Name "QMEnable" -Value 0
        Set-RegistryValue -Path "$officePath\Common\ClientTelemetry" -Name "DisableTelemetry" -Value 1

        # Disable Office updates
        Set-RegistryValue -Path "$officePath\Common\OfficeUpdate" -Name "EnableAutomaticUpdates" -Value 0

        Write-Host "  [+] Office configured" -ForegroundColor Green
        break
    }
}

#endregion

#region Set DNS Servers

Write-Section "Configuring DNS Servers"

Write-Host "Setting DNS servers to 8.8.8.8, 8.8.4.4..." -ForegroundColor Green

# Get active network adapters
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

foreach ($adapter in $adapters) {
    try {
        Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses ("8.8.8.8", "8.8.4.4")
        Write-Host "  [+] DNS configured for $($adapter.Name)" -ForegroundColor Green
    } catch {
        Write-Host "  Warning: Could not set DNS for $($adapter.Name)" -ForegroundColor Yellow
    }
}

#endregion

#region Cleanup

Write-Section "Cleanup"

Write-Host "Cleaning up temporary files..." -ForegroundColor Green

# Clean Chocolatey cache
& choco cache clean --yes 2>&1 | Out-Null

# Clean Windows temp
Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "  [+] Cleanup complete" -ForegroundColor Green

#endregion

#region Summary

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Installed components:" -ForegroundColor Green
Write-Host "  [+] Python 3.8 (32-bit) with Pillow, pywin32" -ForegroundColor White
Write-Host "  [+] .NET Framework (4.8, 4.5.2, 3.5)" -ForegroundColor White
Write-Host "  [+] Java Runtime (OpenJDK 8, 11, Latest)" -ForegroundColor White

if (-not $SkipDevelopmentTools) {
    Write-Host "  [+] Development tools (Git, Node.js, Ruby, Go, PHP)" -ForegroundColor White
}

if (-not $SkipBrowsers) {
    Write-Host "  [+] Web browsers (Chrome, Firefox)" -ForegroundColor White
}

if (-not $SkipAdobeReader) {
    Write-Host "  [+] Adobe Reader DC" -ForegroundColor White
}

if (-not $SkipArchivers) {
    Write-Host "  [+] Archive tools (7-Zip, WinRAR)" -ForegroundColor White
}

if (-not $SkipMediaPlayers) {
    Write-Host "  [+] Media players (VLC, K-Lite)" -ForegroundColor White
}

Write-Host "  [+] Additional tools (Notepad++, PuTTY, WinSCP, etc.)" -ForegroundColor White
Write-Host ""
Write-Host "Security features disabled:" -ForegroundColor Yellow
Write-Host "  * Windows Defender disabled" -ForegroundColor White
Write-Host "  * Windows Firewall disabled" -ForegroundColor White
Write-Host "  * UAC disabled" -ForegroundColor White
Write-Host "  * Windows Update disabled" -ForegroundColor White
Write-Host "  * SmartScreen disabled" -ForegroundColor White
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Install CAPE agent as Windows Service (see guide)" -ForegroundColor White
Write-Host "  2. Install Microsoft Office manually (if needed)" -ForegroundColor White
Write-Host "  3. Test all applications" -ForegroundColor White
Write-Host "  4. Run Windows Update if needed" -ForegroundColor White
Write-Host "  5. Run sysprep when ready" -ForegroundColor White
Write-Host ""
Write-Host "WARNING: This VM is now vulnerable! Use only in isolated network!" -ForegroundColor Red
Write-Host ""

#endregion

# Refresh environment variables
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host ""
