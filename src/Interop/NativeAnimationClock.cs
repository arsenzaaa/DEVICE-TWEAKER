using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace DeviceTweakerCS;

// Unnamed process-local timing objects; no timer-resolution or display changes.
internal static class NativeAnimationClock
{
    [StructLayout(LayoutKind.Explicit, Size = 220)]
    private struct DisplayMode
    {
        [FieldOffset(68)] internal ushort Size;
        [FieldOffset(184)] internal int Frequency;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "EnumDisplaySettingsW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ReadDisplayMode(string deviceName, int mode, ref DisplayMode settings);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateWaitableTimerExW(IntPtr attributes, string? name, uint flags, uint access);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWaitableTimer(SafeWaitHandle timer, ref long dueTime, int period,
        IntPtr completion, IntPtr argument, [MarshalAs(UnmanagedType.Bool)] bool resume);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool CancelWaitableTimer(SafeWaitHandle timer);

    internal static int ReadRefreshRate(Control owner)
    {
        DisplayMode mode = new() { Size = 220 };
        return ReadDisplayMode(Screen.FromControl(owner).DeviceName, -1, ref mode) && mode.Frequency is >= 30 and <= 1000
            ? mode.Frequency : 60;
    }

    internal sealed class FrameTimer : WaitHandle
    {
        internal bool HighResolution { get; }
        internal FrameTimer()
        {
            const uint access = 0x00100002; // SYNCHRONIZE | TIMER_MODIFY_STATE
            IntPtr handle = CreateWaitableTimerExW(IntPtr.Zero, null, 2, access);
            HighResolution = handle != IntPtr.Zero;
            if (handle == IntPtr.Zero) handle = CreateWaitableTimerExW(IntPtr.Zero, null, 0, access);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            SafeWaitHandle = new SafeWaitHandle(handle, ownsHandle: true);
        }

        internal void Arm(double milliseconds)
        {
            long relative = -Math.Max(1, (long)Math.Ceiling(milliseconds * 10000));
            if (!SetWaitableTimer(SafeWaitHandle, ref relative, 0, IntPtr.Zero, IntPtr.Zero, false))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }
}
