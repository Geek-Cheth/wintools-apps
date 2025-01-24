# RAM Optimizer Script v3.2
# Requires administrative privileges to run

# Enable strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Add Windows Forms for system tray icon
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create system tray icon
$sysTrayIcon = New-Object System.Windows.Forms.NotifyIcon
$sysTrayIcon.Text = "RAM Optimizer"
$sysTrayIcon.Icon = [System.Drawing.SystemIcons]::Application
$sysTrayIcon.Visible = $true

# Create context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$toolStripForceOptimize = New-Object System.Windows.Forms.ToolStripMenuItem
$toolStripForceOptimize.Text = "Force Optimization"
$toolStripExit = New-Object System.Windows.Forms.ToolStripMenuItem
$toolStripExit.Text = "Exit"
$contextMenu.Items.Add($toolStripForceOptimize)
$contextMenu.Items.Add($toolStripExit)
$sysTrayIcon.ContextMenuStrip = $contextMenu

# Function to show balloon tip
function Show-Notification {
    param (
        [string]$Title,
        [string]$Message,
        [System.Windows.Forms.ToolTipIcon]$Icon = [System.Windows.Forms.ToolTipIcon]::Info
    )
    $sysTrayIcon.BalloonTipTitle = $Title
    $sysTrayIcon.BalloonTipText = $Message
    $sysTrayIcon.BalloonTipIcon = $Icon
    $sysTrayIcon.ShowBalloonTip(5000)
}

# Check for Admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires Administrator privileges. Please run as Administrator." -ForegroundColor Red
    Exit 1
}

# Import required assemblies
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class MemoryManagement {
        [DllImport("kernel32.dll")]
        public static extern bool SetProcessWorkingSetSize(IntPtr proc, int min, int max);
        
        [DllImport("psapi.dll")]
        public static extern bool EmptyWorkingSet(IntPtr hProcess);
        
        [DllImport("kernel32.dll")]
        public static extern bool SetSystemFileCacheSize(int MinimumFileCacheSize, int MaximumFileCacheSize, int Flags);
    }
"@

# Function to get detailed system memory information
function Get-MemoryDetails {
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        
        $totalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        $freeRAM = [math]::Round($os.FreePhysicalMemory / 1KB / 1024, 2)
        $usedRAM = [math]::Round($totalRAM - $freeRAM, 2)
        $ramUsage = [math]::Round(($usedRAM / $totalRAM) * 100, 2)
        
        return @{
            TotalRAM = $totalRAM
            FreeRAM = $freeRAM
            UsedRAM = $usedRAM
            UsagePercent = $ramUsage
            PageFileUsage = [math]::Round(($os.SizeStoredInPagingFiles - $os.FreeSpaceInPagingFiles) / 1MB, 2)
            CommittedMemory = [math]::Round($os.TotalVirtualMemorySize / 1MB, 2)
        }
    } catch {
        Write-Host "Error getting memory details: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Function to get top memory-consuming processes with more details
function Get-TopMemoryProcesses {
    return Get-Process | 
        Where-Object { -not [string]::IsNullOrEmpty($_.ProcessName) } |
        Sort-Object WorkingSet64 -Descending | 
        Select-Object -First 8 Name, 
            @{Name='PID';Expression={$_.Id}},
            @{Name='Memory(MB)';Expression={[math]::Round($_.WorkingSet64 / 1MB, 2)}},
            @{Name='Private(MB)';Expression={[math]::Round($_.PrivateMemorySize64 / 1MB, 2)}},
            @{Name='CPU(s)';Expression={[math]::Round($_.CPU, 2)}},
            @{Name='ThreadCount';Expression={$_.Threads.Count}},
            @{Name='Handles';Expression={$_.HandleCount}}
}

# Function to optimize specific process memory
function Optimize-ProcessMemory {
    param (
        [Parameter(Mandatory=$true)]
        [int]$processId,
        [switch]$Aggressive
    )
    try {
        $process = Get-Process -Id $processId -ErrorAction Stop
        if ($process.ProcessName -notin @('System', 'Idle', 'Registry')) {
            [MemoryManagement]::EmptyWorkingSet($process.Handle) | Out-Null
            if ($Aggressive) {
                [MemoryManagement]::SetProcessWorkingSetSize($process.Handle, -1, -1) | Out-Null
            }
            return $true
        }
    } catch {
        return $false
    }
}

# Function to clear system caches
function Clear-SystemCaches {
    param([switch]$Aggressive)
    
    try {
        # Clear file system cache
        Write-Host "Clearing file system cache..."
        [MemoryManagement]::SetSystemFileCacheSize(0, 0, 0) | Out-Null
        
        # Clear DNS cache
        Write-Host "Clearing DNS cache..."
        Start-Process "ipconfig" -ArgumentList "/flushdns" -WindowStyle Hidden -Wait
        
        if ($Aggressive) {
            # Clear additional caches
            Write-Host "Clearing additional system caches..."
            $commands = @(
                "netsh interface ip delete arpcache",
                "netsh winsock reset",
                "sc stop sysmain",  # Disable Superfetch
                "sc config sysmain start= disabled"
            )
            foreach ($cmd in $commands) {
                Start-Process "cmd.exe" -ArgumentList "/c $cmd" -WindowStyle Hidden -Wait
            }
        }
        
        # Clear temp files
        $tempPaths = @(
            "$env:TEMP",
            "$env:SystemRoot\Temp",
            "$env:USERPROFILE\AppData\Local\Temp",
            "$env:SystemRoot\Prefetch",
            "$env:USERPROFILE\AppData\Local\Microsoft\Windows\INetCache"
        )
        
        foreach ($path in $tempPaths) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -File -Force | 
                    Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                    Remove-Item -Force -ErrorAction SilentlyContinue
            }
        }
        
        return $true
    } catch {
        Write-Host "Error clearing system caches: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to optimize RAM
function Optimize-RAM {
    param ([switch]$Aggressive)
    
    $startTime = Get-Date
    Write-Host "`n[$(Get-Date)] Starting RAM optimization..." -ForegroundColor Yellow
    Show-Notification "RAM Optimization" "Starting memory optimization..."
    
    try {
        # Get initial memory state
        $initialMemory = Get-MemoryDetails
        
        # Display system memory status
        Write-Host "`nSystem Memory Status:" -ForegroundColor Cyan
        Write-Host "Total RAM: $($initialMemory.TotalRAM) GB"
        Write-Host "Used RAM: $($initialMemory.UsedRAM) GB"
        Write-Host "Free RAM: $($initialMemory.FreeRAM) GB"
        Write-Host "Page File Usage: $($initialMemory.PageFileUsage) MB"
        
        # Display top memory-consuming processes
        Write-Host "`nTop memory-consuming processes before optimization:" -ForegroundColor Cyan
        Get-TopMemoryProcesses | Format-Table -AutoSize

        # Optimize processes
        Write-Host "Optimizing process memory..."
        $optimizedCount = 0
        Get-Process | Where-Object { $_.WorkingSet64 -gt 100MB } | ForEach-Object {
            if (Optimize-ProcessMemory -processId $_.Id -Aggressive:$Aggressive) {
                Write-Host "Optimized: $($_.ProcessName) (PID: $($_.Id))" -ForegroundColor DarkGray
                $optimizedCount++
            }
        }
        
        # Clear system caches
        Clear-SystemCaches -Aggressive:$Aggressive
        
        # Run garbage collection
        [System.GC]::Collect(2, [System.GCCollectionMode]::Forced, $true, $true)
        [System.GC]::WaitForPendingFinalizers()
        
        # Get final memory state
        Start-Sleep -Seconds 2  # Wait for changes to take effect
        $finalMemory = Get-MemoryDetails
        $freedMemory = $initialMemory.UsedRAM - $finalMemory.UsedRAM
        $improvement = $initialMemory.UsagePercent - $finalMemory.UsagePercent
        $duration = (Get-Date) - $startTime
        
        # Display results with notification
        $resultMessage = "Memory freed: $($freedMemory.ToString('0.00')) GB`nImprovement: $($improvement.ToString('0.00'))%"
        Show-Notification "Optimization Complete" $resultMessage
        
        Write-Host "`nOptimization Results:" -ForegroundColor Green
        Write-Host "Duration: $($duration.TotalSeconds.ToString('0.00')) seconds"
        Write-Host "Processes Optimized: $optimizedCount"
        Write-Host "Initial RAM Usage: $($initialMemory.UsagePercent)%" -ForegroundColor Yellow
        Write-Host "Final RAM Usage: $($finalMemory.UsagePercent)%" -ForegroundColor Yellow
        Write-Host "Memory Freed: $($freedMemory.ToString('0.00')) GB" -ForegroundColor Green
        Write-Host "Improvement: $($improvement.ToString('0.00'))%" -ForegroundColor Green
        
        Write-Host "`nCurrent top memory-consuming processes:" -ForegroundColor Cyan
        Get-TopMemoryProcesses | Format-Table -AutoSize
        
    } catch {
        $errorMsg = "Error during optimization: $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
        Show-Notification "Optimization Error" $errorMsg ([System.Windows.Forms.ToolTipIcon]::Error)
    }
    
    Write-Host "`nRAM optimization cycle completed!" -ForegroundColor Green
    Write-Host "----------------------------------------`n"
}

# Configuration
$config = @{
    ThresholdNormal = 70    # Normal threshold (%)
    ThresholdHigh = 85      # High threshold for aggressive mode (%)
    CheckInterval = 30      # Check every 30 seconds
    CooldownNormal = 300    # 5 minutes cooldown for normal optimization
    CooldownAggressive = 600 # 10 minutes cooldown for aggressive optimization
    LogFile = "$env:USERPROFILE\Documents\RAMOptimizer.log"
    MaxLogSize = 10MB       # Maximum log file size
}

# Initialize logging
if (-not (Test-Path $config.LogFile)) {
    New-Item -Path $config.LogFile -ItemType File -Force | Out-Null
}

# Rotate log if too large
if ((Get-Item $config.LogFile).Length -gt $config.MaxLogSize) {
    Move-Item -Path $config.LogFile -Destination "$($config.LogFile).old" -Force
}

$lastOptimizationTime = [DateTime]::MinValue
$script:lastRAMUsage = 0

# Display startup information
Write-Host "RAM Optimizer v3.2 Started" -ForegroundColor Cyan
Write-Host "----------------------------------------"
Write-Host "Normal Threshold: $($config.ThresholdNormal)%"
Write-Host "Aggressive Threshold: $($config.ThresholdHigh)%"
Write-Host "Check Interval: $($config.CheckInterval) seconds"
Write-Host "Normal Cooldown: $($config.CooldownNormal / 60) minutes"
Write-Host "Aggressive Cooldown: $($config.CooldownAggressive / 60) minutes"
Write-Host "Log File: $($config.LogFile)"
Write-Host "----------------------------------------`n"

# Register event handlers
$toolStripForceOptimize.Add_Click({
    Write-Host "Manual optimization requested by user"
    Optimize-RAM -Aggressive
})

$toolStripExit.Add_Click({
    $sysTrayIcon.Visible = $false
    $sysTrayIcon.Dispose()
    Stop-Process $pid
})

# Show startup notification
Show-Notification "RAM Optimizer Started" "Monitoring RAM usage. Normal threshold: $($config.ThresholdNormal)%"

# Main monitoring loop with UI message pump
[System.Windows.Forms.Application]::EnableVisualStyles()
$appContext = New-Object System.Windows.Forms.ApplicationContext

while ($true) {
    try {
        [System.Windows.Forms.Application]::DoEvents()
        
        $memInfo = Get-MemoryDetails
        $currentRAMUsage = $memInfo.UsagePercent
        $script:lastRAMUsage = $currentRAMUsage
        $timeSinceLastOptimization = (Get-Date) - $lastOptimizationTime
        
        # Update system tray icon tooltip
        $sysTrayIcon.Text = "RAM Usage: $($currentRAMUsage)%`nFree: $($memInfo.FreeRAM) GB"
        
        # Log current status
        $status = "RAM Usage: $currentRAMUsage% | Free: $($memInfo.FreeRAM) GB | Used: $($memInfo.UsedRAM) GB | $(Get-Date -Format 'HH:mm:ss')"
        Write-Host $status -ForegroundColor Gray
        Add-Content -Path $config.LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $status"
        
        # Check if optimization is needed
        if ($currentRAMUsage -ge $config.ThresholdHigh -and $timeSinceLastOptimization.TotalSeconds -ge $config.CooldownAggressive) {
            Write-Host "Critical RAM usage detected! ($currentRAMUsage%)" -ForegroundColor Red
            Show-Notification "Critical RAM Usage" "RAM usage is at $currentRAMUsage%. Starting aggressive optimization..." ([System.Windows.Forms.ToolTipIcon]::Warning)
            Optimize-RAM -Aggressive
            $lastOptimizationTime = Get-Date
        }
        elseif ($currentRAMUsage -ge $config.ThresholdNormal -and $timeSinceLastOptimization.TotalSeconds -ge $config.CooldownNormal) {
            Write-Host "High RAM usage detected! ($currentRAMUsage%)" -ForegroundColor Yellow
            Show-Notification "High RAM Usage" "RAM usage is at $currentRAMUsage%. Starting optimization..."
            Optimize-RAM
            $lastOptimizationTime = Get-Date
        }
        
        Start-Sleep -Seconds $config.CheckInterval
        
    } catch {
        $errorMessage = "Error in main loop: $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Add-Content -Path $config.LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): ERROR - $errorMessage"
        Show-Notification "Error" $errorMessage ([System.Windows.Forms.ToolTipIcon]::Error)
        Start-Sleep -Seconds 10
    }
} 