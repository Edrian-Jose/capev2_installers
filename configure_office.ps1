<#
.SYNOPSIS
    Configure Microsoft Office for malware analysis

.DESCRIPTION
    Configures Microsoft Office to disable security features, updates,
    and telemetry for malware analysis environment. Run AFTER installing Office.

.NOTES
    - Run as Administrator
    - Run AFTER installing Microsoft Office
    - Works with Office 2013, 2016, 2019, 365

.EXAMPLE
    .\configure_office.ps1
#>

[CmdletBinding()]
param()

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Microsoft Office Configuration for CAPE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

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
        Write-Host "  Warning: Could not set $Path\$Name" -ForegroundColor Yellow
        return $false
    }
}

# Detect Office version
$officeVersions = @("16.0", "15.0", "14.0")  # 2016/2019/365, 2013, 2010
$installedVersion = $null

foreach ($version in $officeVersions) {
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\$version") {
        $installedVersion = $version
        break
    }
}

if (-not $installedVersion) {
    Write-Host "Microsoft Office not detected!" -ForegroundColor Red
    Write-Host "Please install Office before running this script." -ForegroundColor Yellow
    Exit 1
}

Write-Host "Detected Office version: $installedVersion" -ForegroundColor Green
Write-Host ""

$basePath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\$installedVersion"

# Office applications to configure
$apps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher")

Write-Host "Configuring Office applications..." -ForegroundColor Yellow
Write-Host ""

foreach ($app in $apps) {
    Write-Host "Configuring $app..." -ForegroundColor Green

    # Disable macro warnings (enable all macros - DANGEROUS!)
    Set-RegistryValue -Path "$basePath\$app\Security" -Name "VBAWarnings" -Value 1
    Set-RegistryValue -Path "$basePath\$app\Security" -Name "AccessVBOM" -Value 1

    # Disable Protected View for all sources
    Set-RegistryValue -Path "$basePath\$app\Security\ProtectedView" -Name "DisableAttachmentsInPV" -Value 1
    Set-RegistryValue -Path "$basePath\$app\Security\ProtectedView" -Name "DisableInternetFilesInPV" -Value 1
    Set-RegistryValue -Path "$basePath\$app\Security\ProtectedView" -Name "DisableUnsafeLocationsInPV" -Value 1

    # Disable Trust Bar notifications
    Set-RegistryValue -Path "$basePath\$app\Security" -Name "NotbPrmssns" -Value 1

    # Enable ActiveX without prompting (DANGEROUS!)
    Set-RegistryValue -Path "$basePath\$app\Security" -Name "DisableAllActiveX" -Value 0
    Set-RegistryValue -Path "$basePath\$app\Security" -Name "UFIControls" -Value 1

    Write-Host "  ✓ $app configured" -ForegroundColor Green
}

Write-Host ""
Write-Host "Configuring Office Common settings..." -ForegroundColor Green

# Disable Office telemetry
Set-RegistryValue -Path "$basePath\Common" -Name "QMEnable" -Value 0
Set-RegistryValue -Path "$basePath\Common\ClientTelemetry" -Name "DisableTelemetry" -Value 1
Set-RegistryValue -Path "$basePath\Common\ClientTelemetry" -Name "SendTelemetry" -Value 3

# Disable Office updates
Set-RegistryValue -Path "$basePath\Common\OfficeUpdate" -Name "EnableAutomaticUpdates" -Value 0
Set-RegistryValue -Path "$basePath\Common\OfficeUpdate" -Name "HideEnableDisableUpdates" -Value 1

# Disable customer experience improvement program
Set-RegistryValue -Path "$basePath\Common" -Name "UpdateReliabilityData" -Value 0

# Disable first run prompts
Set-RegistryValue -Path "$basePath\Common\General" -Name "ShownFirstRunOptin" -Value 1
Set-RegistryValue -Path "$basePath\Common" -Name "OptInDisable" -Value 1

# Disable privacy dialog
Set-RegistryValue -Path "$basePath\Common\Privacy" -Name "DisconnectedState" -Value 2
Set-RegistryValue -Path "$basePath\Common\Privacy" -Name "UserContentDisabled" -Value 2

# Disable feedback and surveys
Set-RegistryValue -Path "$basePath\Common\Feedback" -Name "Enabled" -Value 0
Set-RegistryValue -Path "$basePath\Common\Feedback" -Name "IncludeEmail" -Value 0

# Disable Office Intelligent Services
Set-RegistryValue -Path "$basePath\Common\Privacy" -Name "ControllerConnectedServicesEnabled" -Value 2
Set-RegistryValue -Path "$basePath\Common\Privacy" -Name "DownloadContentDisabled" -Value 2

Write-Host "  ✓ Office Common configured" -ForegroundColor Green
Write-Host ""

# Configure Trust Center
Write-Host "Configuring Trust Center..." -ForegroundColor Green

foreach ($app in $apps) {
    # Disable all security warnings
    Set-RegistryValue -Path "$basePath\$app\Security\Trusted Locations\Location0" -Name "AllowSubfolders" -Value 1
    Set-RegistryValue -Path "$basePath\$app\Security\Trusted Locations\Location0" -Name "Path" -Value "C:\" -Type "String"

    # Disable file validation
    Set-RegistryValue -Path "$basePath\$app\Security\FileValidation" -Name "EnableOnLoad" -Value 0
    Set-RegistryValue -Path "$basePath\$app\Security\FileValidation" -Name "DisableEditFromPV" -Value 0

    # Allow legacy file formats
    Set-RegistryValue -Path "$basePath\$app\Options" -Name "DontUpdateLinks" -Value 0
}

Write-Host "  ✓ Trust Center configured" -ForegroundColor Green
Write-Host ""

# Disable Office Click-to-Run service updates
Write-Host "Configuring Office Click-to-Run..." -ForegroundColor Green

$c2rPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate"
Set-RegistryValue -Path $c2rPath -Name "EnableAutomaticUpdates" -Value 0
Set-RegistryValue -Path $c2rPath -Name "HideEnableDisableUpdates" -Value 1
Set-RegistryValue -Path $c2rPath -Name "HideUpdateNotifications" -Value 1

# Disable Office update tasks
$tasks = @(
    "Office Automatic Updates*",
    "Office ClickToRun*",
    "Office Feature Updates*"
)

foreach ($taskPattern in $tasks) {
    Get-ScheduledTask -TaskName $taskPattern -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue
}

Write-Host "  ✓ Office updates disabled" -ForegroundColor Green
Write-Host ""

# Summary
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Configuration Complete!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Office security features disabled:" -ForegroundColor Yellow
Write-Host "  • Macro warnings disabled (all macros enabled)" -ForegroundColor White
Write-Host "  • Protected View disabled" -ForegroundColor White
Write-Host "  • ActiveX controls enabled" -ForegroundColor White
Write-Host "  • File validation disabled" -ForegroundColor White
Write-Host "  • Trust Bar notifications disabled" -ForegroundColor White
Write-Host "  • Office updates disabled" -ForegroundColor White
Write-Host "  • Telemetry and feedback disabled" -ForegroundColor White
Write-Host ""
Write-Host "WARNING: Office is now vulnerable to macro malware!" -ForegroundColor Red
Write-Host "Use only in isolated malware analysis environment!" -ForegroundColor Red
Write-Host ""
Write-Host "These settings will survive sysprep." -ForegroundColor Green
Write-Host ""
