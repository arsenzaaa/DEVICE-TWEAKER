using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal static class NativeUxTheme
{
    internal enum PreferredAppMode
    {
        Default = 0,
        AllowDark = 1,
        ForceDark = 2,
        ForceLight = 3,
        Max = 4,
    }

    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);

    [DllImport("uxtheme.dll", EntryPoint = "#135", SetLastError = true)]
    internal static extern int SetPreferredAppMode(PreferredAppMode appMode);

    [DllImport("uxtheme.dll", EntryPoint = "#133", SetLastError = true)]
    internal static extern int AllowDarkModeForWindow(IntPtr hWnd, bool allow);

    [DllImport("uxtheme.dll", EntryPoint = "#104", SetLastError = true)]
    internal static extern void RefreshImmersiveColorPolicyState();
}
