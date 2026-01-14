<#
.SYNOPSIS
    Install CAPE agent as Windows Service that survives sysprep

.DESCRIPTION
    This script installs the CAPE malware analysis agent (agent.py) as a
    Windows Service running under LocalSystem account. The service auto-starts
    on boot and survives sysprep/generalization.

    Requirements:
    - Python 3.8 32-bit installed at C:\Python38-32\
    - agent.py files in Downloads folder or specify -AgentSourcePath
    - Run as Administrator

.PARAMETER AgentSourcePath
    Path to agent.py folder. Default: $env:USERPROFILE\Downloads\agent

.PARAMETER SkipPythonPackages
    Skip installing Python packages (Pillow, pywin32)

.EXAMPLE
    .\install_cape_agent.ps1

.EXAMPLE
    .\install_cape_agent.ps1 -AgentSourcePath "C:\Temp\agent"

.EXAMPLE
    .\install_cape_agent.ps1 -SkipPythonPackages
#>

[CmdletBinding()]
param(
    [string]$AgentSourcePath = "$env:USERPROFILE\Downloads\agent",
    [switch]$SkipPythonPackages
)

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    Exit 1
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "CAPE Agent Installation for Windows" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$pythonExe = "C:\Python38-32\python.exe"
$pythonScripts = "C:\Python38-32\Scripts"
$agentDestPath = "C:\CAPE\agent"
$serviceWrapperPath = "C:\CAPE\agent_service.py"
$serviceName = "CAPEAgent"
$serviceDisplayName = "CAPE Malware Analysis Agent"
$agentPort = 8000

#region Step 1: Verify Python Installation

Write-Host "Step 1: Verifying Python installation..." -ForegroundColor Yellow
Write-Host ""

if (-not (Test-Path $pythonExe)) {
    Write-Host "[-] Python 3.8 32-bit not found at: $pythonExe" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python 3.8 32-bit first:" -ForegroundColor Yellow
    Write-Host "  1. Run: .\install_analysis_tools.ps1" -ForegroundColor White
    Write-Host "  OR" -ForegroundColor White
    Write-Host "  2. Download from: https://www.python.org/downloads/" -ForegroundColor White
    Write-Host "     Install to: C:\Python38-32\" -ForegroundColor White
    Write-Host ""
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

    Write-Host "Installing pywin32 (required for Windows Service)..." -ForegroundColor Cyan
    & $pythonExe -m pip install pywin32 --quiet

    Write-Host "[+] Python packages installed successfully" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "Step 2: Skipping Python package installation" -ForegroundColor Yellow
    Write-Host ""
}

#endregion

#region Step 3: Copy Agent Files

Write-Host "Step 3: Copying agent files..." -ForegroundColor Yellow
Write-Host ""

# Check if agent source exists
if (-not (Test-Path $AgentSourcePath)) {
    Write-Host "[-] Agent source path not found: $AgentSourcePath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please ensure agent.py is available in one of these locations:" -ForegroundColor Yellow
    Write-Host "  - $env:USERPROFILE\Downloads\agent\" -ForegroundColor White
    Write-Host "  - Specify custom path with -AgentSourcePath parameter" -ForegroundColor White
    Write-Host ""
    Exit 1
}

# Check if agent.py exists
$agentPySource = Join-Path $AgentSourcePath "agent.py"
if (-not (Test-Path $agentPySource)) {
    Write-Host "[-] agent.py not found in: $AgentSourcePath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Expected file: $agentPySource" -ForegroundColor Yellow
    Write-Host ""
    Exit 1
}

Write-Host "[+] Found agent.py at: $agentPySource" -ForegroundColor Green

# Create destination directory
if (-not (Test-Path $agentDestPath)) {
    New-Item -ItemType Directory -Path $agentDestPath -Force | Out-Null
    Write-Host "[+] Created directory: $agentDestPath" -ForegroundColor Green
}

# Copy all agent files
Write-Host "Copying agent files..." -ForegroundColor Cyan
try {
    Copy-Item -Path "$AgentSourcePath\*" -Destination $agentDestPath -Recurse -Force
    Write-Host "[+] Agent files copied successfully" -ForegroundColor Green
    Write-Host "    Destination: $agentDestPath" -ForegroundColor DarkGray
} catch {
    Write-Host "[-] Failed to copy agent files: $_" -ForegroundColor Red
    Exit 1
}

Write-Host ""

#endregion

#region Step 4: Create Windows Service Wrapper

Write-Host "Step 4: Creating Windows Service wrapper..." -ForegroundColor Yellow
Write-Host ""

$serviceWrapperContent = @'
"""
CAPE Agent Windows Service Wrapper
This script runs agent.py as a Windows Service that survives sysprep

Service Configuration:
- Runs under LocalSystem account
- Auto-starts on boot
- Auto-restarts on failure
- Waits 30 seconds for network initialization
- Monitors and restarts agent.py if it crashes
"""
import sys
import os
import time
import subprocess
import servicemanager
import win32serviceutil
import win32service
import win32event

class CAPEAgentService(win32serviceutil.ServiceFramework):
    """Windows Service wrapper for CAPE agent.py"""

    _svc_name_ = "CAPEAgent"
    _svc_display_name_ = "CAPE Malware Analysis Agent"
    _svc_description_ = "Cuckoo/CAPE guest agent for malware analysis. Listens on port 8000."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.agent_process = None
        self.running = True

    def SvcStop(self):
        """Called when service is stopped"""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.stop_event)

        # Terminate agent process
        if self.agent_process:
            try:
                self.agent_process.terminate()
                self.agent_process.wait(timeout=10)
            except:
                pass

    def SvcDoRun(self):
        """Called when service starts"""
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.main()

    def main(self):
        """Main service loop"""
        # Wait for network to be ready (critical for Azure VMs)
        servicemanager.LogInfoMsg("Waiting 30 seconds for network initialization...")
        time.sleep(30)

        agent_path = r"C:\CAPE\agent\agent.py"
        python_exe = r"C:\Python38-32\python.exe"

        # Verify files exist
        if not os.path.exists(agent_path):
            servicemanager.LogErrorMsg(f"Agent not found: {agent_path}")
            return

        if not os.path.exists(python_exe):
            servicemanager.LogErrorMsg(f"Python not found: {python_exe}")
            return

        servicemanager.LogInfoMsg(f"Starting CAPE agent: {agent_path}")

        # Main loop - keep agent running
        restart_count = 0
        max_restarts_per_minute = 5

        while self.running:
            try:
                # Start agent.py
                servicemanager.LogInfoMsg("Starting agent.py process...")

                self.agent_process = subprocess.Popen(
                    [python_exe, agent_path],
                    cwd=r"C:\CAPE\agent",
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                servicemanager.LogInfoMsg(f"Agent started with PID: {self.agent_process.pid}")
                restart_count = 0

                # Monitor agent process
                while self.running:
                    # Check if service stop was requested
                    rc = win32event.WaitForSingleObject(self.stop_event, 5000)
                    if rc == win32event.WAIT_OBJECT_0:
                        # Service stop requested
                        servicemanager.LogInfoMsg("Service stop requested")
                        return

                    # Check if agent process is still running
                    if self.agent_process.poll() is not None:
                        # Agent exited
                        exit_code = self.agent_process.returncode
                        servicemanager.LogWarningMsg(f"Agent process exited with code: {exit_code}")

                        # Wait before restart
                        time.sleep(5)
                        break

            except Exception as e:
                servicemanager.LogErrorMsg(f"Agent error: {str(e)}")
                restart_count += 1

                # Prevent restart loop
                if restart_count >= max_restarts_per_minute:
                    servicemanager.LogErrorMsg(f"Too many restarts ({restart_count}), waiting 60 seconds...")
                    time.sleep(60)
                    restart_count = 0
                else:
                    time.sleep(10)

        servicemanager.LogInfoMsg("Service stopped")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        # Called by Windows Service Manager
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(CAPEAgentService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Called from command line for install/remove/debug
        win32serviceutil.HandleCommandLine(CAPEAgentService)
'@

try {
    $serviceWrapperContent | Out-File -FilePath $serviceWrapperPath -Encoding UTF8 -Force
    Write-Host "[+] Created service wrapper: $serviceWrapperPath" -ForegroundColor Green
} catch {
    Write-Host "[-] Failed to create service wrapper: $_" -ForegroundColor Red
    Exit 1
}

Write-Host ""

#endregion

#region Step 5: Install Windows Service

Write-Host "Step 5: Installing Windows Service..." -ForegroundColor Yellow
Write-Host ""

# Remove existing service if present
$existingService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Removing existing service..." -ForegroundColor Cyan

    # Stop service if running
    if ($existingService.Status -eq 'Running') {
        Stop-Service -Name $serviceName -Force
        Start-Sleep -Seconds 2
    }

    # Remove service
    & $pythonExe $serviceWrapperPath remove
    Start-Sleep -Seconds 2
    Write-Host "[+] Existing service removed" -ForegroundColor Green
}

# Install service
Write-Host "Installing CAPE Agent service..." -ForegroundColor Cyan
try {
    & $pythonExe $serviceWrapperPath install

    if ($LASTEXITCODE -eq 0) {
        Write-Host "[+] Service installed successfully" -ForegroundColor Green
    } else {
        Write-Host "[-] Service installation failed with exit code: $LASTEXITCODE" -ForegroundColor Red
        Exit 1
    }
} catch {
    Write-Host "[-] Failed to install service: $_" -ForegroundColor Red
    Exit 1
}

Write-Host ""

#endregion

#region Step 6: Configure Service

Write-Host "Step 6: Configuring service..." -ForegroundColor Yellow
Write-Host ""

# Set service to auto-start
Write-Host "Setting auto-start..." -ForegroundColor Cyan
sc.exe config $serviceName start= auto | Out-Null

# Set service to run as LocalSystem
Write-Host "Setting LocalSystem account..." -ForegroundColor Cyan
sc.exe config $serviceName obj= LocalSystem | Out-Null

# Set service recovery options (auto-restart on failure)
Write-Host "Configuring failure recovery..." -ForegroundColor Cyan
sc.exe failure $serviceName reset= 86400 actions= restart/5000/restart/10000/restart/30000 | Out-Null

# Set service description
sc.exe description $serviceName "CAPE/Cuckoo malware analysis agent. Runs agent.py on port 8000. Required for sandbox analysis." | Out-Null

Write-Host "[+] Service configured successfully" -ForegroundColor Green
Write-Host "    - Auto-start: Enabled" -ForegroundColor DarkGray
Write-Host "    - Account: LocalSystem" -ForegroundColor DarkGray
Write-Host "    - Recovery: Auto-restart on failure" -ForegroundColor DarkGray
Write-Host ""

#endregion

#region Step 7: Start Service

Write-Host "Step 7: Starting service..." -ForegroundColor Yellow
Write-Host ""

try {
    Start-Service -Name $serviceName
    Start-Sleep -Seconds 5

    $service = Get-Service -Name $serviceName
    if ($service.Status -eq 'Running') {
        Write-Host "[+] Service started successfully" -ForegroundColor Green
        Write-Host "    Status: $($service.Status)" -ForegroundColor DarkGray
    } else {
        Write-Host "[-] Service status: $($service.Status)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Check service logs:" -ForegroundColor Yellow
        Write-Host "  Event Viewer > Windows Logs > Application" -ForegroundColor White
        Write-Host "  Look for 'CAPEAgent' events" -ForegroundColor White
        Exit 1
    }
} catch {
    Write-Host "[-] Failed to start service: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Check Event Viewer > Application logs" -ForegroundColor White
    Write-Host "  2. Verify Python path: $pythonExe" -ForegroundColor White
    Write-Host "  3. Verify agent path: $agentDestPath\agent.py" -ForegroundColor White
    Write-Host "  4. Check pywin32 is installed: $pythonExe -m pip show pywin32" -ForegroundColor White
    Exit 1
}

Write-Host ""

#endregion

#region Step 8: Verify Agent is Listening

Write-Host "Step 8: Verifying agent is listening on port $agentPort..." -ForegroundColor Yellow
Write-Host ""

# Wait for agent to start listening
Write-Host "Waiting for agent to initialize (30 seconds)..." -ForegroundColor Cyan
Start-Sleep -Seconds 30

# Check if port 8000 is listening
$listening = Get-NetTCPConnection -LocalPort $agentPort -State Listen -ErrorAction SilentlyContinue

if ($listening) {
    Write-Host "[+] Agent is listening on port $agentPort" -ForegroundColor Green
    Write-Host "    Local Address: $($listening.LocalAddress):$($listening.LocalPort)" -ForegroundColor DarkGray
    Write-Host "    Process: $($listening.OwningProcess)" -ForegroundColor DarkGray
} else {
    Write-Host "[-] Agent is NOT listening on port $agentPort" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Check service status: sc query $serviceName" -ForegroundColor White
    Write-Host "  2. Check service logs in Event Viewer" -ForegroundColor White
    Write-Host "  3. Check Windows Firewall settings" -ForegroundColor White
    Write-Host "  4. Manually test agent:" -ForegroundColor White
    Write-Host "     $pythonExe $agentDestPath\agent.py" -ForegroundColor White
    Write-Host ""

    # Show recent Python processes
    $pythonProcs = Get-Process -Name python -ErrorAction SilentlyContinue
    if ($pythonProcs) {
        Write-Host "Running Python processes:" -ForegroundColor Cyan
        $pythonProcs | Format-Table -Property Id, ProcessName, Path -AutoSize
    }
}

Write-Host ""

#endregion

#region Step 9: Create Firewall Rule

Write-Host "Step 9: Creating Windows Firewall rule..." -ForegroundColor Yellow
Write-Host ""

try {
    # Remove existing rule if present
    Remove-NetFirewallRule -DisplayName "CAPE Agent" -ErrorAction SilentlyContinue | Out-Null

    # Create inbound rule for port 8000
    New-NetFirewallRule -DisplayName "CAPE Agent" `
        -Description "Allow inbound connections to CAPE malware analysis agent on port 8000" `
        -Direction Inbound `
        -Protocol TCP `
        -LocalPort $agentPort `
        -Action Allow `
        -Profile Any `
        -Enabled True | Out-Null

    Write-Host "[+] Firewall rule created (allows port $agentPort)" -ForegroundColor Green
} catch {
    Write-Host "[!] Warning: Could not create firewall rule: $_" -ForegroundColor Yellow
    Write-Host "    You may need to manually allow port $agentPort" -ForegroundColor Yellow
}

Write-Host ""

#endregion

#region Summary

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$service = Get-Service -Name $serviceName
$serviceConfig = sc.exe qc $serviceName

Write-Host "Service Information:" -ForegroundColor Green
Write-Host "  Name: $serviceName" -ForegroundColor White
Write-Host "  Display Name: $serviceDisplayName" -ForegroundColor White
Write-Host "  Status: $($service.Status)" -ForegroundColor White
Write-Host "  Startup Type: Automatic" -ForegroundColor White
Write-Host "  Account: LocalSystem" -ForegroundColor White
Write-Host ""

Write-Host "Agent Information:" -ForegroundColor Green
Write-Host "  Location: $agentDestPath\agent.py" -ForegroundColor White
Write-Host "  Python: $pythonExe" -ForegroundColor White
Write-Host "  Port: $agentPort" -ForegroundColor White
Write-Host "  Listening: $(if ($listening) { 'YES' } else { 'NO' })" -ForegroundColor $(if ($listening) { 'Green' } else { 'Red' })
Write-Host ""

Write-Host "Features:" -ForegroundColor Green
Write-Host "  [+] Auto-starts on boot" -ForegroundColor White
Write-Host "  [+] Runs under LocalSystem (survives sysprep)" -ForegroundColor White
Write-Host "  [+] Auto-restarts on failure" -ForegroundColor White
Write-Host "  [+] 30-second network initialization delay" -ForegroundColor White
Write-Host "  [+] Firewall rule configured" -ForegroundColor White
Write-Host "  [+] All files in C:\CAPE\ (survives sysprep)" -ForegroundColor White
Write-Host ""

Write-Host "Useful Commands:" -ForegroundColor Cyan
Write-Host "  Check status:  sc query $serviceName" -ForegroundColor White
Write-Host "  Start service: net start $serviceName" -ForegroundColor White
Write-Host "  Stop service:  net stop $serviceName" -ForegroundColor White
Write-Host "  View logs:     Event Viewer > Application > CAPEAgent" -ForegroundColor White
Write-Host "  Test manually: $pythonExe $agentDestPath\agent.py" -ForegroundColor White
Write-Host "  Check port:    netstat -ano | findstr :$agentPort" -ForegroundColor White
Write-Host ""

Write-Host "Testing from CAPE host:" -ForegroundColor Cyan
Write-Host "  telnet <vm-ip> $agentPort" -ForegroundColor White
Write-Host "  curl http://<vm-ip>:$agentPort/status" -ForegroundColor White
Write-Host ""

if ($listening) {
    Write-Host "SUCCESS: CAPE Agent is running and ready!" -ForegroundColor Green
    Write-Host "You can now run sysprep - the service will survive!" -ForegroundColor Green
} else {
    Write-Host "WARNING: Agent may not be fully initialized yet." -ForegroundColor Yellow
    Write-Host "Wait 1-2 minutes and check: netstat -ano | findstr :$agentPort" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test agent is accessible from CAPE host" -ForegroundColor White
Write-Host "  2. Verify all applications are installed" -ForegroundColor White
Write-Host "  3. Run: .\generate_dummy_files.ps1" -ForegroundColor White
Write-Host "  4. Test everything works" -ForegroundColor White
Write-Host "  5. Run sysprep when ready" -ForegroundColor White
Write-Host ""

#endregion
