#To test manually
#powershell -ExecutionPolicy Bypass -File ".\downloadCounter.ps1"

```powershell
# ============================================
# Downloads & Recycle Bin Counter
# Counts files in Downloads and items in the
# Recycle Bin, then shows a Windows notification
# when you log in.
# ============================================

$downloads = Join-Path $env:USERPROFILE "Downloads"

# ============================================
# Count Downloads
# ============================================

$countDownloads = (Get-ChildItem -Path $downloads -File -ErrorAction SilentlyContinue).Count

# ============================================
# Count Recycle Bin
# ============================================

$shell = New-Object -ComObject Shell.Application
$recycleBin = $shell.Namespace(10)

$countRecycleBin = $recycleBin.Items().Count

# ============================================
# Create startup shortcut
# ============================================

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "Downloads & Recycle Bin Counter.lnk"
$scriptPath = $MyInvocation.MyCommand.Path

# Only create the shortcut if it doesn't already exist
if (-not (Test-Path $shortcutPath)) {

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutPath)

    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    $shortcut.WorkingDirectory = Split-Path $scriptPath
    $shortcut.Save()
}

# ============================================
# Windows notification
# ============================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$notification = New-Object System.Windows.Forms.NotifyIcon
$notification.Icon = [System.Drawing.SystemIcons]::Information
$notification.Visible = $true

$notification.BalloonTipTitle = "Downloads & Recycle Bin"
$notification.BalloonTipText = @"
Downloads: $countDownloads files
Recycle Bin: $countRecycleBin items
"@

$notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info

$notification.ShowBalloonTip(5000)

# Keep the notification alive long enough to display
Start-Sleep -Seconds 6

$notification.Dispose()
```
