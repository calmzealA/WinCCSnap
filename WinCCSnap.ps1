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
    [ValidateSet('install', 'info', 'remove', 'help')]
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
'@ | Out-File -FilePath $Script -Encoding UTF8 -Force

    $act = New-ScheduledTaskAction -Execute "powershell.exe" `
               -Argument "-NoProfile -WindowStyle Hidden -File `"$Script`""
    $trg = New-ScheduledTaskTrigger -AtStartup
    $set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
               -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask -TaskName $TaskName -Action $act -Trigger $trg `
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
        Write-Info "WinCCSnap installed successfully. Background listener started (Job ID: $($job.Id))" Green
    } catch {
        Write-Info "WinCCSnap installed successfully. Reboot or run the task to enable." Yellow
    }
}

function Show-Status {
    $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($t) {
        Write-Info "WinCCSnap is installed. Current state: $($t.State)" Green
    } else {
        Write-Info "WinCCSnap is not installed." Yellow
    }
}

function Remove-WinCCSnap {
    if (-not (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
        Write-Info "WinCCSnap is not installed – nothing to remove." Yellow
        return
    }
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Remove-Item -Path $ScriptDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Info "WinCCSnap has been uninstalled." Green
}

function Show-Help {
    Write-Info "WinCCSnap – One-time clipboard helper for Claude Code on Windows"
    Write-Info "Usage:"
    Write-Info "  .\CCClip.ps1 install   # Setup once (autostart)"
    Write-Info "  .\CCClip.ps1 info      # Show current status"
    Write-Info "  .\CCClip.ps1 remove    # Uninstall completely"
    Write-Info "  .\CCClip.ps1 help      # Show this help"
}

switch ($Action) {
    'install' { Install-WinCCSnap }
    'info'    { Show-Status }
    'remove'  { Remove-WinCCSnap }
    'help'    { Show-Help }
}