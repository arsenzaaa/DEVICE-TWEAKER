param(
    [ValidateSet('GpuOnly', 'ReadmeOnly')][string]$Mode = 'GpuOnly',
    [string]$OutputDirectory = (Join-Path $env:TEMP 'DeviceTweakerShowcaseHidden')
)

$ErrorActionPreference = 'Stop'
if (-not ([System.Management.Automation.PSTypeName]'ShowcaseDesktopNative').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class ShowcaseDesktopNative {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STARTUPINFO {
        public int cb;
        public string lpReserved;
        public string lpDesktop;
        public string lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars;
        public int dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2;
        public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION {
        public IntPtr hProcess, hThread;
        public int dwProcessId, dwThreadId;
    }
    [DllImport("user32.dll", EntryPoint = "CreateDesktopW", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr devmode, uint flags, uint access, IntPtr security);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool CreateProcess(string application, StringBuilder commandLine, IntPtr processAttributes,
        IntPtr threadAttributes, bool inheritHandles, uint creationFlags, IntPtr environment, string currentDirectory,
        ref STARTUPINFO startup, out PROCESS_INFORMATION info);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetExitCodeProcess(IntPtr handle, out uint exitCode);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr handle);
    public static Exception LastError(string call) {
        int code = Marshal.GetLastWin32Error();
        return new Win32Exception(code, call + " failed with Win32 error " + code);
    }
}
'@
}

$destination = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $destination | Out-Null
$statusPath = Join-Path $destination 'capture.log'
$worker = Join-Path $PSScriptRoot 'Capture-Showcase.HiddenWorker.ps1'
$hostExe = (Get-Process -Id $PID).Path
$desktopName = 'DeviceTweakerShowcase_' + [Guid]::NewGuid().ToString('N')
$desktop = [ShowcaseDesktopNative]::CreateDesktop($desktopName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x000001FF, [IntPtr]::Zero)
if ($desktop -eq [IntPtr]::Zero) { throw [ShowcaseDesktopNative]::LastError('CreateDesktop') }

$info = New-Object ShowcaseDesktopNative+PROCESS_INFORMATION
try {
    $startup = New-Object ShowcaseDesktopNative+STARTUPINFO
    $startup.cb = [Runtime.InteropServices.Marshal]::SizeOf([type]'ShowcaseDesktopNative+STARTUPINFO')
    $startup.lpDesktop = $desktopName
    $command = '"' + $hostExe + '" -NoProfile -ExecutionPolicy Bypass -File "' + $worker + '" -Mode ' + $Mode + ' -OutputDirectory "' + $destination + '" -StatusPath "' + $statusPath + '"'
    $commandLine = [System.Text.StringBuilder]::new($command)
    $ok = [ShowcaseDesktopNative]::CreateProcess($hostExe, $commandLine, [IntPtr]::Zero, [IntPtr]::Zero,
        $false, 0x08000000, [IntPtr]::Zero, $PSScriptRoot, [ref]$startup, [ref]$info)
    if (-not $ok) { throw [ShowcaseDesktopNative]::LastError('CreateProcess') }
    $waitResult = [ShowcaseDesktopNative]::WaitForSingleObject($info.hProcess, 180000)
    if ($waitResult -ne 0) { throw "Hidden showcase worker did not finish (wait result $waitResult, PID $($info.dwProcessId))." }
    [uint32]$exitCode = 0
    if (-not [ShowcaseDesktopNative]::GetExitCodeProcess($info.hProcess, [ref]$exitCode)) {
        throw [ShowcaseDesktopNative]::LastError('GetExitCodeProcess')
    }
    if (Test-Path -LiteralPath $statusPath) { Get-Content -LiteralPath $statusPath }
    if ($exitCode -ne 0) { throw "Hidden showcase worker exited with code $exitCode." }
    Write-Host "Hidden desktop capture complete: $destination"
} finally {
    if ($info.hThread -ne [IntPtr]::Zero) { [ShowcaseDesktopNative]::CloseHandle($info.hThread) | Out-Null }
    if ($info.hProcess -ne [IntPtr]::Zero) { [ShowcaseDesktopNative]::CloseHandle($info.hProcess) | Out-Null }
    [ShowcaseDesktopNative]::CloseDesktop($desktop) | Out-Null
}
