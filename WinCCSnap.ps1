<#
.SYNOPSIS
  WinCCSnap – Automatically convert clipboard bitmap → PNG
  so you can paste images directly into Claude Code on Windows.

.PARAMETER Action
  install : One-time setup (scheduled task + autostart)
  info    : Show current installation status
  remove  : Uninstall everything
  help    : Display this help
#>
param(
    [Parameter(Position = 0)]
    [ValidateSet('install', 'info', 'remove', 'help', 'test', 'restart', 'debug')]
    [string]$Action = 'help'
)

$TaskName  = "WinCCSnapListener"
$ScriptDir = "$env:APPDATA\WinCCSnap"
$Script    = "$ScriptDir\listener.ps1"

function Write-Info([string]$msg, [ConsoleColor]$color = 'Cyan') {
    Write-Host $msg -ForegroundColor $color
}

function Install-WinCCSnap {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Write-Info "WinCCSnap is already installed." Green
        return
    }

    if (-not (Test-Path $ScriptDir)) { New-Item -ItemType Directory -Path $ScriptDir | Out-Null }

    @'
# WinCCSnap Listener with Debug Logging
$logFile = "$env:TEMP\WinCCSnap.log"
$clipFile = "$env:TEMP\clip.png"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

Write-Log "WinCCSnap listener started" "START"
Write-Log "Log file: $logFile"
Write-Log "Clip file: $clipFile"

Add-Type -AssemblyName System.Windows.Forms

$loopCount = 0
$processedCount = 0

while ($true) {
    $loopCount++
    
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            Write-Log "Found image in clipboard (Loop $loopCount)" "PROCESS"
            
            try {
                $img = [System.Windows.Forms.Clipboard]::GetImage()
                Write-Log "Image retrieved: $($img.Width)x$($img.Height)" "PROCESS"
                
                if ($img.Width -gt 0 -and $img.Height -gt 0) {
                    $img.Save($clipFile, [System.Drawing.Imaging.ImageFormat]::Png)
                    Write-Log "Image saved to: $clipFile" "SUCCESS"
                    
                    Set-Clipboard -Path $clipFile
                    Write-Log "PNG file path set to clipboard" "SUCCESS"
                    
                    $processedCount++
                    Write-Log "Total processed: $processedCount" "STATS"
                } else {
                    Write-Log "Invalid image dimensions: $($img.Width)x$($img.Height)" "ERROR"
                }
            } catch {
                Write-Log "Error processing image: $($_.Exception.Message)" "ERROR"
            }
        } else {
            if ($loopCount % 100 -eq 0) {
                Write-Log "No image found (Loop $loopCount)" "DEBUG"
            }
        }
    } catch {
        Write-Log "Error checking clipboard: $($_.Exception.Message)" "ERROR"
    }
    
    Start-Sleep -Milliseconds 500
}
'@ | Out-File -FilePath $Script -Encoding UTF8 -Force

    $act = New-ScheduledTaskAction -Execute "powershell.exe" `
               -Argument "-NoProfile -WindowStyle Hidden -File `"$Script`""
    $trg1 = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $trg2 = New-ScheduledTaskTrigger -AtStartup
    $set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
               -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask -TaskName $TaskName -Action $act -Trigger @($trg1, $trg2) `
        -Settings $set -RunLevel Highest -User $env:USERNAME -Force | Out-Null

    
    # 立即在当前会话中启动监听器
    try {
        $job = Start-Job -ScriptBlock {
            Add-Type -AssemblyName System.Windows.Forms
            while ($true) {
                if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                    $img  = [System.Windows.Forms.Clipboard]::GetImage()
                    $file = "$env:TEMP\clip.png"
                    $img.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
                    Set-Clipboard -Path $file
                    Start-Sleep -Milliseconds 500
                }
                Start-Sleep -Milliseconds 500
            }
        } -Name "WinCCSnapListener"
        Write-Info "WinCCSnap installed successfully with dual triggers (startup + logon)" Green
        Write-Info "Background listener started (Job ID: $($job.Id)) - Ready for immediate use" Green
        Write-Info "Note: Scheduled task will start automatically on next login/reboot" Yellow
    } catch {
        Write-Info "WinCCSnap installed successfully. Reboot or run 'restart' to enable." Yellow
    }
}

function Show-Status {
    Write-Info "=== WinCCSnap Status Check ===" Cyan
    
    # 检查Scheduled Task
    $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($t) {
        Write-Info "📋 Scheduled Task: INSTALLED" Green
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Info "   State: $($t.State)" 
        Write-Info "   Last Run Time: $($taskInfo.LastRunTime)"
        Write-Info "   Next Run Time: $($taskInfo.NextRunTime)"
        
        # 检查日志文件
        $logFile = "$env:TEMP\WinCCSnap.log"
        if (Test-Path $logFile) {
            Write-Info "📄 Log File: EXISTS" Green
            $logContent = Get-Content $logFile -Tail 5
            Write-Info "   Recent logs:"
            $logContent | ForEach-Object { Write-Info "     $_" Gray }
        } else {
            Write-Info "📄 Log File: NOT FOUND" Yellow
        }
        
        # 检查临时文件
        $clipFile = "$env:TEMP\clip.png"
        if (Test-Path $clipFile) {
            $fileInfo = Get-Item $clipFile
            Write-Info "🖼️  Clip File: EXISTS (Modified: $($fileInfo.LastWriteTime))" Green
        } else {
            Write-Info "🖼️  Clip File: NOT FOUND" Yellow
        }
    } else {
        Write-Info "📋 Scheduled Task: NOT INSTALLED" Red
    }
    
    # 检查当前会话的Job
    $job = Get-Job -Name "WinCCSnapListener" -ErrorAction SilentlyContinue
    if ($job) {
        Write-Info "🔄 Current Session Job: RUNNING (ID: $($job.Id))" Green
        Write-Info "   Job State: $($job.State)"
    } else {
        Write-Info "🔄 Current Session Job: NOT RUNNING" Yellow
        Write-Info "   (This is normal after restart - use 'restart' command to start immediately)"
    }
    
    # 检查监听器文件
    if (Test-Path $Script) {
        Write-Info "📁 Listener Script: EXISTS" Green
    } else {
        Write-Info "📁 Listener Script: MISSING" Red
    }
    
    Write-Info ""
    Write-Info "💡 Debug Commands:"
    Write-Info "   Get-Content '$env:TEMP\WinCCSnap.log' -Tail 20" Blue
    Write-Info "   Start-ScheduledTask -TaskName '$TaskName'"
}

function Remove-WinCCSnap {
    Write-Info "=== Removing WinCCSnap ===" Cyan
    
    # 停止并移除所有相关的PowerShell Job
    $jobs = Get-Job -Name "WinCCSnapListener" -ErrorAction SilentlyContinue
    if ($jobs) {
        Write-Info "Stopping running jobs..." Yellow
        $jobs | Stop-Job -Confirm:$false
        $jobs | Remove-Job -Confirm:$false
        Write-Info "✅ Jobs stopped and removed" Green
    }
    
    # 移除计划任务
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Info "✅ Scheduled task removed" Green
    } else {
        Write-Info "⚠️  Scheduled task not found" Yellow
    }
    
    # 移除监听器脚本
    if (Test-Path $ScriptDir) {
        Remove-Item -Path $ScriptDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Info "✅ Listener script removed" Green
    } else {
        Write-Info "⚠️  Listener script directory not found" Yellow
    }
    
    Write-Info "WinCCSnap has been completely uninstalled." Green
}

function Start-WinCCSnapJob {
    try {
        $existingJob = Get-Job -Name "WinCCSnapListener" -ErrorAction SilentlyContinue
        if ($existingJob) {
            Write-Info "Job is already running (ID: $($existingJob.Id))" Yellow
            return
        }
        
        $job = Start-Job -ScriptBlock {
            Add-Type -AssemblyName System.Windows.Forms
            while ($true) {
                if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                    $img  = [System.Windows.Forms.Clipboard]::GetImage()
                    $file = "$env:TEMP\clip.png"
                    $img.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
                    Set-Clipboard -Path $file
                    Start-Sleep -Milliseconds 500
                }
                Start-Sleep -Milliseconds 500
            }
        } -Name "WinCCSnapListener"
        Write-Info "WinCCSnap background listener started (Job ID: $($job.Id))" Green
    } catch {
        Write-Info "Failed to start background listener: $($_.Exception.Message)" Red
    }
}

function Test-WinCCSnap {
    Write-Info "=== Testing WinCCSnap Functionality ===" Cyan
    
    # 检查剪贴板监听器状态
    Show-Status
    
    Write-Info ""
    Write-Info "📋 Test Steps:"
    Write-Info "1. Take a screenshot (Win+Shift+S)"
    Write-Info "2. Check if %TEMP%\clip.png is created"
    Write-Info "3. Try pasting into Claude Code"
    Write-Info ""
    
    # 检查当前剪贴板内容
    try {
        Add-Type -AssemblyName System.Windows.Forms
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            Write-Info "✅ Current clipboard contains image data" Green
        } else {
            Write-Info "⚠️  Current clipboard does not contain image data" Yellow
        }
    } catch {
        Write-Info "❌ Cannot access clipboard: $($_.Exception.Message)" Red
    }
    
    # 检查临时文件
    $tempFile = "$env:TEMP\clip.png"
    if (Test-Path $tempFile) {
        $fileInfo = Get-Item $tempFile
        Write-Info "📁 Found existing clip.png (Modified: $($fileInfo.LastWriteTime))" Green
    } else {
        Write-Info "📁 No clip.png found yet (normal if no images processed)" Yellow
    }
    
    # 测试手动运行计划任务
    Write-Info ""
    Write-Info "🔧 Testing scheduled task manually..." Blue
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) {
        Write-Info "Starting scheduled task manually..." Yellow
        try {
            Start-ScheduledTask -TaskName $TaskName
            Start-Sleep -Seconds 2
            $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
            Write-Info "Task started successfully! Last run: $($taskInfo.LastRunTime)" Green
        } catch {
            Write-Info "Failed to start task: $($_.Exception.Message)" Red
        }
    }
}

function Show-DebugInfo {
    Write-Info "=== WinCCSnap Debug Information ===" Cyan
    
    $logFile = "$env:TEMP\WinCCSnap.log"
    $clipFile = "$env:TEMP\clip.png"
    
    Write-Info "📄 Log File Location: $logFile" Blue
    Write-Info "🖼️  Clip File Location: $clipFile" Blue
    
    # 显示完整日志
    if (Test-Path $logFile) {
        Write-Info ""
        Write-Info "📋 Recent Logs:" Green
        Get-Content $logFile -Tail 10 | ForEach-Object { Write-Info "  $_" }
    } else {
        Write-Info "📋 No logs found - task may not have run yet" Yellow
    }
    
    # 显示系统信息
    Write-Info ""
    Write-Info "🔧 System Information:" Blue
    Write-Info "  PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Info "  Windows Version: $([System.Environment]::OSVersion.Version)"
    Write-Info "  Current User: $env:USERNAME"
    Write-Info "  TEMP Path: $env:TEMP"
    
    # 检查剪贴板权限
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $canAccess = [System.Windows.Forms.Clipboard]::ContainsImage()
        Write-Info "✅ Clipboard access: OK" Green
    } catch {
        Write-Info "❌ Clipboard access error: $($_.Exception.Message)" Red
    }
    
    Write-Info ""
    Write-Info "💡 Manual Debug Steps:" Yellow
    Write-Info "  1. Take screenshot (Win+Shift+S)"
    Write-Info "  2. Check log: Get-Content '$logFile' -Tail 20"
    Write-Info "  3. Check file: Test-Path '$clipFile'"
    Write-Info "  4. Start task: Start-ScheduledTask -TaskName '$TaskName'"
}

function Show-Help {
    Write-Info "WinCCSnap – One-time clipboard helper for Claude Code on Windows"
    Write-Info "Usage:"
    Write-Info "  .\WinCCSnap.ps1 install   # Setup once (scheduled task + autostart)"
    Write-Info "  .\WinCCSnap.ps1 info      # Show detailed status (task + job + logs)"
    Write-Info "  .\WinCCSnap.ps1 remove    # Uninstall completely"
    Write-Info "  .\WinCCSnap.ps1 restart   # Start listener in current session"
    Write-Info "  .\WinCCSnap.ps1 test      # Test functionality"
    Write-Info "  .\WinCCSnap.ps1 debug     # Show debug information"
    Write-Info "  .\WinCCSnap.ps1 help      # Show this help"
    Write-Info ""
    Write-Info "💡 Debug Commands:"
    Write-Info "   Get-Content '$env:TEMP\WinCCSnap.log' -Tail 50"
    Write-Info "   Start-ScheduledTask -TaskName 'WinCCSnapListener'"
}

switch ($Action) {
    'install' { Install-WinCCSnap }
    'info'    { Show-Status }
    'remove'  { Remove-WinCCSnap }
    'test'    { Test-WinCCSnap }
    'restart' { Start-WinCCSnapJob }
    'debug'   { Show-DebugInfo }
    'help'    { Show-Help }
}