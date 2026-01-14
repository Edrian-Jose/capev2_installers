<#
.SYNOPSIS
    Install CAPE agent for SPECIALIZED Azure image (no sysprep)

.DESCRIPTION
    This script installs the CAPE malware analysis agent (agent.py) with
    simple auto-start configuration that works with SPECIALIZED Azure images.

    No Windows Service needed - uses simple Task Scheduler task that survives
    in SPECIALIZED images.

    The script auto-detects agent.py in common locations:
    - Downloads\agent.py (file directly in Downloads)
    - Downloads\agent\ (folder with agent.py inside)
    - Desktop\agent.py or Desktop\agent\
    - C:\Temp\agent.py or C:\Temp\agent\

    Requirements:
    - Python 3.8 32-bit installed at C:\Python38-32\
    - agent.py file (auto-detected or specify -AgentSourcePath)
    - Run as Administrator

.PARAMETER AgentSourcePath
    Path to agent.py file or folder containing agent.py.
    If not specified, auto-detects in common locations.

.PARAMETER SkipPythonPackages
    Skip installing Python packages (Pillow)

.EXAMPLE
    .\install_cape_agent_specialized.ps1

.EXAMPLE
    .\install_cape_agent_specialized.ps1 -AgentSourcePath "C:\Users\Admin\Downloads\agent.py"
#>

[CmdletBinding()]
param(
    [string]$AgentSourcePath = "",
    [switch]$SkipPythonPackages
)

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "CAPE Agent Installation (SPECIALIZED Image)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Note: This setup is for SPECIALIZED Azure images (no sysprep)" -ForegroundColor Yellow
Write-Host ""

# Configuration
$pythonExe = "C:\Python38-32\python.exe"
$agentDestPath = "C:\CAPE\agent"
$agentPort = 8000

#region Step 1: Verify Python Installation

Write-Host "Step 1: Verifying Python installation..." -ForegroundColor Yellow
Write-Host ""

if (-not (Test-Path $pythonExe)) {
    Write-Host "[-] Python 3.8 32-bit not found at: $pythonExe" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python 3.8 32-bit first:" -ForegroundColor Yellow
    Write-Host "  Run: .\install_analysis_tools.ps1" -ForegroundColor White
    Exit 1
}

$pythonVersion = & $pythonExe --version 2>&1
Write-Host "[+] Found: $pythonVersion" -ForegroundColor Green
Write-Host "    Location: $pythonExe" -ForegroundColor DarkGray
Write-Host ""

#endregion

#region Step 2: Install Python Packages

if (-not $SkipPythonPackages) {
    Write-Host "Step 2: Installing Python packages..." -ForegroundColor Yellow
    Write-Host ""

    Write-Host "Installing pip..." -ForegroundColor Cyan
    & $pythonExe -m pip install --upgrade pip --quiet

    Write-Host "Installing Pillow..." -ForegroundColor Cyan
    & $pythonExe -m pip install Pillow --quiet

    Write-Host "[+] Python packages installed successfully" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "Step 2: Skipping Python package installation" -ForegroundColor Yellow
    Write-Host ""
}

#endregion

#region Step 3: Locate and Copy Agent Files

Write-Host "Step 3: Locating and copying agent files..." -ForegroundColor Yellow
Write-Host ""

# Auto-detect agent.py location if not specified
if ([string]::IsNullOrEmpty($AgentSourcePath)) {
    Write-Host "Auto-detecting agent.py location..." -ForegroundColor Cyan

    $commonLocations = @(
        "$env:USERPROFILE\Downloads\agent.py",
        "$env:USERPROFILE\Downloads\agent",
        "$env:USERPROFILE\Desktop\agent.py",
        "$env:USERPROFILE\Desktop\agent",
        "C:\Temp\agent.py",
        "C:\Temp\agent"
    )

    $found = $false
    foreach ($location in $commonLocations) {
        if (Test-Path $location) {
            $item = Get-Item $location
            if ($item.PSIsContainer) {
                $agentPyPath = Join-Path $location "agent.py"
                if (Test-Path $agentPyPath) {
                    $AgentSourcePath = $location
                    $found = $true
                    Write-Host "[+] Found agent.py in folder: $AgentSourcePath" -ForegroundColor Green
                    break
                }
            } else {
                if ($item.Name -eq "agent.py") {
                    $AgentSourcePath = $item.DirectoryName
                    $found = $true
                    Write-Host "[+] Found agent.py file at: $location" -ForegroundColor Green
                    break
                }
            }
        }
    }

    if (-not $found) {
        Write-Host "[-] Could not auto-detect agent.py location" -ForegroundColor Red
        Write-Host ""
        Write-Host "Please specify location with -AgentSourcePath parameter" -ForegroundColor Yellow
        Exit 1
    }
}

# Normalize path
$agentSourceItem = Get-Item $AgentSourcePath -ErrorAction SilentlyContinue
if (-not $agentSourceItem) {
    Write-Host "[-] Path not found: $AgentSourcePath" -ForegroundColor Red
    Exit 1
}

if (-not $agentSourceItem.PSIsContainer -and $agentSourceItem.Name -eq "agent.py") {
    $AgentSourcePath = $agentSourceItem.DirectoryName
    Write-Host "[+] Using directory: $AgentSourcePath" -ForegroundColor Green
}

# Verify agent.py exists
$agentPySource = Join-Path $AgentSourcePath "agent.py"
if (-not (Test-Path $agentPySource)) {
    Write-Host "[-] agent.py not found in: $AgentSourcePath" -ForegroundColor Red
    Exit 1
}

Write-Host "[+] Found agent.py at: $agentPySource" -ForegroundColor Green

# Create destination and copy
if (-not (Test-Path $agentDestPath)) {
    New-Item -ItemType Directory -Path $agentDestPath -Force | Out-Null
    Write-Host "[+] Created directory: $agentDestPath" -ForegroundColor Green
}

Write-Host "Copying agent files..." -ForegroundColor Cyan
Copy-Item -Path "$AgentSourcePath\*" -Destination $agentDestPath -Recurse -Force
Write-Host "[+] Agent files copied successfully" -ForegroundColor Green
Write-Host ""

#endregion

#region Step 4: Create Auto-Start Task

Write-Host "Step 4: Creating auto-start configuration..." -ForegroundColor Yellow
Write-Host ""

# Create Task Scheduler task (runs under SYSTEM, survives SPECIALIZED image)
Write-Host "Creating Task Scheduler task..." -ForegroundColor Cyan

# Remove existing task if present
Unregister-ScheduledTask -TaskName "CAPEAgent" -Confirm:$false -ErrorAction SilentlyContinue

# Create task action
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument "$agentDestPath\agent.py" -WorkingDirectory $agentDestPath

# Create task trigger (at startup)
$trigger = New-ScheduledTaskTrigger -AtStartup

# Create task principal (SYSTEM account)
$principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Create task settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0)

# Register task
Register-ScheduledTask `
    -TaskName "CAPEAgent" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "CAPE malware analysis agent. Starts agent.py on boot." | Out-Null

# Add 30-second delay for network initialization
$task = Get-ScheduledTask -TaskName "CAPEAgent"
$task.Triggers[0].Delay = "PT30S"
$task | Set-ScheduledTask | Out-Null

Write-Host "[+] Task Scheduler task created" -ForegroundColor Green
Write-Host "    Task Name: CAPEAgent" -ForegroundColor DarkGray
Write-Host "    Trigger: At startup (30s delay)" -ForegroundColor DarkGray
Write-Host "    Account: SYSTEM" -ForegroundColor DarkGray
Write-Host ""

#endregion

#region Step 5: Start Agent Manually

Write-Host "Step 5: Starting agent..." -ForegroundColor Yellow
Write-Host ""

# Start task now to test
Write-Host "Starting agent task..." -ForegroundColor Cyan
Start-ScheduledTask -TaskName "CAPEAgent"

Start-Sleep -Seconds 5

# Check if running
$pythonProcs = Get-Process -Name python -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $pythonExe }
if ($pythonProcs) {
    Write-Host "[+] Agent process is running" -ForegroundColor Green
    Write-Host "    PID: $($pythonProcs[0].Id)" -ForegroundColor DarkGray
} else {
    Write-Host "[-] Agent process not found" -ForegroundColor Red
    Write-Host "    Check task manually: Get-ScheduledTask -TaskName CAPEAgent" -ForegroundColor Yellow
}

Write-Host ""

#endregion

#region Step 6: Verify Agent is Listening

Write-Host "Step 6: Verifying agent is listening..." -ForegroundColor Yellow
Write-Host ""

Write-Host "Waiting for agent to initialize (10 seconds)..." -ForegroundColor Cyan
Start-Sleep -Seconds 10

$listening = Get-NetTCPConnection -LocalPort $agentPort -State Listen -ErrorAction SilentlyContinue

if ($listening) {
    Write-Host "[+] Agent is listening on port $agentPort" -ForegroundColor Green
    Write-Host "    Local Address: $($listening.LocalAddress):$($listening.LocalPort)" -ForegroundColor DarkGray
} else {
    Write-Host "[-] Agent is NOT listening on port $agentPort" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Check task: Get-ScheduledTask -TaskName CAPEAgent" -ForegroundColor White
    Write-Host "  2. Check process: Get-Process python" -ForegroundColor White
    Write-Host "  3. Manually test: $pythonExe $agentDestPath\agent.py" -ForegroundColor White
}

Write-Host ""

#endregion

#region Step 7: Create Firewall Rule

Write-Host "Step 7: Creating Windows Firewall rule..." -ForegroundColor Yellow
Write-Host ""

try {
    Remove-NetFirewallRule -DisplayName "CAPE Agent" -ErrorAction SilentlyContinue | Out-Null

    New-NetFirewallRule -DisplayName "CAPE Agent" `
        -Description "Allow inbound connections to CAPE agent on port 8000" `
        -Direction Inbound `
        -Protocol TCP `
        -LocalPort $agentPort `
        -Action Allow `
        -Profile Any `
        -Enabled True | Out-Null

    Write-Host "[+] Firewall rule created (allows port $agentPort)" -ForegroundColor Green
} catch {
    Write-Host "[!] Warning: Could not create firewall rule" -ForegroundColor Yellow
}

Write-Host ""

#endregion

#region Summary

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$task = Get-ScheduledTask -TaskName "CAPEAgent"

Write-Host "Agent Information:" -ForegroundColor Green
Write-Host "  Location: $agentDestPath\agent.py" -ForegroundColor White
Write-Host "  Python: $pythonExe" -ForegroundColor White
Write-Host "  Port: $agentPort" -ForegroundColor White
Write-Host "  Listening: $(if ($listening) { 'YES' } else { 'NO' })" -ForegroundColor $(if ($listening) { 'Green' } else { 'Red' })
Write-Host ""

Write-Host "Auto-Start Configuration:" -ForegroundColor Green
Write-Host "  Task Name: CAPEAgent" -ForegroundColor White
Write-Host "  Status: $($task.State)" -ForegroundColor White
Write-Host "  Account: SYSTEM" -ForegroundColor White
Write-Host "  Trigger: At startup (30s delay)" -ForegroundColor White
Write-Host ""

Write-Host "Features:" -ForegroundColor Green
Write-Host "  [+] Auto-starts on boot" -ForegroundColor White
Write-Host "  [+] Runs under SYSTEM account" -ForegroundColor White
Write-Host "  [+] Auto-restarts on failure" -ForegroundColor White
Write-Host "  [+] 30-second network initialization delay" -ForegroundColor White
Write-Host "  [+] Works with SPECIALIZED Azure images" -ForegroundColor White
Write-Host "  [+] NO sysprep needed!" -ForegroundColor White
Write-Host ""

Write-Host "Useful Commands:" -ForegroundColor Cyan
Write-Host "  Check task:    Get-ScheduledTask -TaskName CAPEAgent" -ForegroundColor White
Write-Host "  Start task:    Start-ScheduledTask -TaskName CAPEAgent" -ForegroundColor White
Write-Host "  Stop task:     Stop-ScheduledTask -TaskName CAPEAgent" -ForegroundColor White
Write-Host "  Check port:    netstat -ano | findstr :$agentPort" -ForegroundColor White
Write-Host "  Test manually: $pythonExe $agentDestPath\agent.py" -ForegroundColor White
Write-Host ""

if ($listening) {
    Write-Host "SUCCESS: CAPE Agent is running and ready!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Test agent from CAPE host: telnet <vm-ip> $agentPort" -ForegroundColor White
    Write-Host "  2. Install other applications" -ForegroundColor White
    Write-Host "  3. Generate dummy files" -ForegroundColor White
    Write-Host "  4. Create SPECIALIZED snapshot (NO sysprep!)" -ForegroundColor White
} else {
    Write-Host "WARNING: Agent may not be fully initialized yet." -ForegroundColor Yellow
    Write-Host "Wait 1-2 minutes and check: netstat -ano | findstr :$agentPort" -ForegroundColor Yellow
}

Write-Host ""

#endregion
