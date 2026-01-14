<#
.SYNOPSIS
    Add installed analysis tools to System PATH

.DESCRIPTION
    This script adds all installed analysis tools to the system PATH environment
    variable so they can be accessed from any command prompt or PowerShell session.
    Runs after install_analysis_tools.ps1 to make all tools globally accessible.

.NOTES
    - Run as Administrator
    - Run AFTER install_analysis_tools.ps1
    - Modifies system PATH (survives sysprep)
    - Does not add duplicates

.EXAMPLE
    .\add_to_system_path.ps1

.EXAMPLE
    .\add_to_system_path.ps1 -Verbose
#>

[CmdletBinding()]
param()

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Add Analysis Tools to System PATH" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

function Add-ToSystemPath {
    <#
    .SYNOPSIS
        Adds a directory to system PATH if it exists and is not already in PATH
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$false)]
        [string]$Description
    )

    # Check if path exists
    if (-not (Test-Path $Path)) {
        Write-Host "[-] Skipping $Description - Path not found: $Path" -ForegroundColor Yellow
        return $false
    }

    # Get current system PATH
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")

    # Check if already in PATH
    $pathElements = $currentPath -split ';' | ForEach-Object { $_.TrimEnd('\') }
    $normalizedNewPath = $Path.TrimEnd('\')

    if ($pathElements -contains $normalizedNewPath) {
        Write-Host "[=] Already in PATH: $Description" -ForegroundColor Gray
        return $false
    }

    # Add to PATH
    try {
        $newPath = $currentPath.TrimEnd(';') + ';' + $normalizedNewPath
        [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
        Write-Host "[+] Added to PATH: $Description" -ForegroundColor Green
        Write-Host "    $Path" -ForegroundColor DarkGray
        return $true
    } catch {
        Write-Host "[-] Failed to add $Description : $_" -ForegroundColor Red
        return $false
    }
}

Write-Host "Adding installed tools to system PATH..." -ForegroundColor Yellow
Write-Host ""

$pathsAdded = 0

# Python 3.8 32-bit
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Python38-32" -Description "Python 3.8 32-bit")
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Python38-32\Scripts" -Description "Python 3.8 Scripts")

# Java
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\OpenJDK\jdk-17\bin" -Description "OpenJDK 17")
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\OpenJDK\jdk-11\bin" -Description "OpenJDK 11")
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\OpenJDK\jdk-8\bin" -Description "OpenJDK 8")

# Also check for alternate Java install locations
if (Test-Path "C:\Program Files\Eclipse Adoptium") {
    $jdkDirs = Get-ChildItem "C:\Program Files\Eclipse Adoptium" -Directory | Where-Object { $_.Name -like "jdk*" }
    foreach ($jdkDir in $jdkDirs) {
        $jdkBin = Join-Path $jdkDir.FullName "bin"
        if (Test-Path $jdkBin) {
            $pathsAdded += [int](Add-ToSystemPath -Path $jdkBin -Description "Java $($jdkDir.Name)")
        }
    }
}

# Git
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Git\cmd" -Description "Git")
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Git\bin" -Description "Git Bin (Unix tools)")

# Node.js
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\nodejs" -Description "Node.js")

# Yarn
$yarnPath = "$env:APPDATA\npm"
$pathsAdded += [int](Add-ToSystemPath -Path $yarnPath -Description "Yarn/NPM Global")

# Ruby
if (Test-Path "C:\tools\ruby32") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\tools\ruby32\bin" -Description "Ruby")
}
if (Test-Path "C:\tools\ruby31") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\tools\ruby31\bin" -Description "Ruby")
}

# Go
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Go\bin" -Description "Go (Golang)")
$goPath = "$env:USERPROFILE\go\bin"
if (Test-Path $goPath) {
    $pathsAdded += [int](Add-ToSystemPath -Path $goPath -Description "Go User Bin")
}

# PHP
if (Test-Path "C:\tools\php83") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\tools\php83" -Description "PHP 8.3")
}
if (Test-Path "C:\tools\php82") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\tools\php82" -Description "PHP 8.2")
}
if (Test-Path "C:\tools\php81") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\tools\php81" -Description "PHP 8.1")
}

# Composer
$pathsAdded += [int](Add-ToSystemPath -Path "C:\ProgramData\ComposerSetup\bin" -Description "Composer")
$composerVendor = "$env:APPDATA\Composer\vendor\bin"
if (Test-Path $composerVendor) {
    $pathsAdded += [int](Add-ToSystemPath -Path $composerVendor -Description "Composer Vendor Bin")
}

# 7-Zip
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\7-Zip" -Description "7-Zip")

# WinRAR
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\WinRAR" -Description "WinRAR")

# VLC
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\VideoLAN\VLC" -Description "VLC Media Player")

# Notepad++
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Notepad++" -Description "Notepad++")

# PuTTY
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\PuTTY" -Description "PuTTY")

# WinSCP
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files (x86)\WinSCP" -Description "WinSCP")

# cURL
$pathsAdded += [int](Add-ToSystemPath -Path "C:\ProgramData\chocolatey\bin" -Description "Chocolatey Bin (curl, wget, etc.)")

# Wget
# Already in chocolatey\bin

# Postman (CLI if available)
$postmanPath = "$env:LOCALAPPDATA\Postman"
if (Test-Path $postmanPath) {
    $pathsAdded += [int](Add-ToSystemPath -Path $postmanPath -Description "Postman")
}

# .NET SDK
if (Test-Path "C:\Program Files\dotnet") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\dotnet" -Description ".NET SDK")
}

# Chrome (for automation/testing)
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Google\Chrome\Application" -Description "Google Chrome")

# Firefox (for automation/testing)
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Mozilla Firefox" -Description "Mozilla Firefox")

# Adobe Reader
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader" -Description "Adobe Reader DC")
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\Adobe\Acrobat DC\Acrobat" -Description "Adobe Acrobat DC")

# LibreOffice
$pathsAdded += [int](Add-ToSystemPath -Path "C:\Program Files\LibreOffice\program" -Description "LibreOffice")

# Microsoft Office (if installed)
$officePaths = @(
    "C:\Program Files\Microsoft Office\root\Office16",
    "C:\Program Files (x86)\Microsoft Office\root\Office16",
    "C:\Program Files\Microsoft Office\Office16",
    "C:\Program Files (x86)\Microsoft Office\Office16"
)

foreach ($officePath in $officePaths) {
    if (Test-Path $officePath) {
        $pathsAdded += [int](Add-ToSystemPath -Path $officePath -Description "Microsoft Office")
        break  # Only add one Office path
    }
}

# CAPE Agent (for manual testing)
if (Test-Path "C:\CAPE\agent") {
    $pathsAdded += [int](Add-ToSystemPath -Path "C:\CAPE\agent" -Description "CAPE Agent")
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

if ($pathsAdded -gt 0) {
    Write-Host "[+] Added $pathsAdded new path(s) to system PATH" -ForegroundColor Green
} else {
    Write-Host "[=] No new paths added (all already in PATH or not found)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "IMPORTANT: Changes to system PATH will take effect:" -ForegroundColor Yellow
Write-Host "  1. Immediately for NEW command prompts/PowerShell windows" -ForegroundColor White
Write-Host "  2. After reboot for existing processes" -ForegroundColor White
Write-Host "  3. For current session, run: RefreshEnv (or restart terminal)" -ForegroundColor White
Write-Host ""

# Optionally refresh for current session
Write-Host "Refreshing environment variables for current session..." -ForegroundColor Yellow
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

Write-Host "[+] Current session PATH updated" -ForegroundColor Green
Write-Host ""
Write-Host "You can now run these tools from any command prompt:" -ForegroundColor Cyan
Write-Host "  - python, pip" -ForegroundColor White
Write-Host "  - java, javac" -ForegroundColor White
Write-Host "  - git, node, npm, yarn" -ForegroundColor White
Write-Host "  - ruby, go, php, composer" -ForegroundColor White
Write-Host "  - 7z, winrar, vlc" -ForegroundColor White
Write-Host "  - curl, wget, notepad++" -ForegroundColor White
Write-Host "  - and more!" -ForegroundColor White
Write-Host ""

# Test a few critical commands
Write-Host "Testing critical commands..." -ForegroundColor Yellow
Write-Host ""

$commands = @(
    @{Name="Python"; Command="python"; Args="--version"},
    @{Name="Git"; Command="git"; Args="--version"},
    @{Name="Node.js"; Command="node"; Args="--version"},
    @{Name="Java"; Command="java"; Args="-version"}
)

foreach ($cmd in $commands) {
    try {
        $output = & $cmd.Command $cmd.Args 2>&1 | Select-Object -First 1
        Write-Host "[+] $($cmd.Name): $output" -ForegroundColor Green
    } catch {
        Write-Host "[-] $($cmd.Name): Not found in PATH" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host "System PATH has been updated and will survive sysprep." -ForegroundColor Green
Write-Host ""
