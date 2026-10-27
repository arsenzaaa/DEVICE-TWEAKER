using System.Runtime.InteropServices;
using System.Text;

namespace DeviceTweakerCS;

internal static class QaDesktop
{
    internal static readonly string? Expected = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_DESKTOP");
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(uint id);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetUserObjectInformationW(IntPtr obj, int index, StringBuilder text, int bytes, out int needed);
    internal static string CurrentName()
    {
        StringBuilder text = new(512);
        if (!GetUserObjectInformationW(GetThreadDesktop(GetCurrentThreadId()), 2, text, text.Capacity * 2, out _))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        return text.ToString();
    }
    internal static void DemandAssignedDesktop()
    {
        if (Expected is null) return;
        if (!QaRun.IsReadOnly || !Expected.StartsWith("DeviceTweakerQa_", StringComparison.Ordinal)
            || !string.Equals(CurrentName(), Expected, StringComparison.Ordinal))
            throw new InvalidOperationException("QA requires the assigned private desktop; no windows were created.");
    }
}
