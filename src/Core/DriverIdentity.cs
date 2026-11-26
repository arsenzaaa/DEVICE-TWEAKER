using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace DeviceTweakerCS;

internal static class DriverIdentity
{
    internal const string ServiceName = DeviceTweakerImodRuntime.ImodDriverServiceName;
    internal const string DevicePath = @"\\.\DeviceTweakerImod2";
    internal const string LegacyDevicePath = @"\\.\DTIMOD";
    internal const string RestartMessage = DeviceTweakerImodRuntime.OlderDriverMessage;

    internal static void CheckLegacyDevice()
    {
        QaRun.DemandWrite("DTIMOD LEGACY DEVICE CHECK");
        CheckLegacyDevice(DeviceAvailable);
    }

    // The current driver cannot safely share its hardware interface with the old build.
    // This preflight only opens the old device; it never changes Windows services.
    internal static void CheckLegacyDevice(Func<string, bool> deviceAvailable)
    {
        if (deviceAvailable(LegacyDevicePath)) throw new InvalidOperationException(RestartMessage);
    }

    internal static bool IsLoaded()
        => DeviceAvailable(DevicePath) || DeviceAvailable(LegacyDevicePath);

    // A service from an earlier build may still reference the staged driver.
    internal static bool HasRegisteredService()
        => IsServiceRegistered(ServiceName) || IsServiceRegistered("DeviceTweakerImod2") || IsServiceRegistered("DTIMOD");

    private static bool IsServiceRegistered(string name)
    {
        DriverServiceState service = DriverEnvironment.ReadService(name);
        return service.Exists ?? throw new InvalidOperationException($"Could not query service {name}: {service.State}.");
    }

    private static bool DeviceAvailable(string path)
    {
        using SafeFileHandle handle = CreateFile(path, 0x80000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (!handle.IsInvalid) return true;
        int error = Marshal.GetLastWin32Error();
        if (error is 2 or 3) return false;
        throw new Win32Exception(error, $"Could not check the DTIMOD device {path}.");
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint share,
        IntPtr security, uint creation, uint flags, IntPtr template);
}
