Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    }
"@

function Start-HiddenOutlook {
    Write-Host "Starting Outlook..."
    Start-Process olk
    
    Write-Host "Waiting for Outlook window to appear..."
    Start-Sleep -Seconds 1

    $enumWindowsCallback = {
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        $className = New-Object System.Text.StringBuilder 256
        [Win32]::GetClassName($hWnd, $className, 256) | Out-Null
        if ($className.ToString() -eq "Olk Host") {
            Write-Host "Found Outlook window: $($hWnd)"
            [Win32]::ShowWindow($hWnd, 0)
        }
        return $true
    }

    $enumWindowsDelegate = [Win32+EnumWindowsProc]$enumWindowsCallback
    [Win32]::EnumWindows($enumWindowsDelegate, [IntPtr]::Zero) | Out-Null
}

function Check-And-Start-Outlook {
    $outlookProcess = Get-Process olk -ErrorAction SilentlyContinue
    if (-not $outlookProcess) {
        Write-Host "Outlook is not running. Starting Outlook..."
        Start-HiddenOutlook
    } else {
        Write-Host "Outlook is already running."
    }
}

Check-And-Start-Outlook
