using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace DeviceTweakerCS;

internal static class PciInterruptReader
{
    private const uint IoctlQueryPci = (0x8010u << 16) | (1u << 14) | (0x814u << 2);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeviceIoControl(IntPtr handle, uint code, byte[] input, int inputSize, byte[] output, int outputSize, out int returned, IntPtr overlapped);

    internal static InterruptModeResult QueryLoaded(string instanceId)
    {
        if (QaRun.IsReadOnly) return InterruptModeResult.Pending;
        if (!instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
            return new(InterruptMode.NotPci, "This device has no PCI interrupt mode.");
        using SafeFileHandle handle = CreateFile(DriverRename.DevicePath, 0x80000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (handle.IsInvalid)
        {
            int error = Marshal.GetLastWin32Error();
            var allocation = WindowsInterruptReader.Read(instanceId);
            return error is 2 or 3 ? allocation : allocation with
            { Detail = $"PCI device open failed ({error}); using Windows allocation. " + allocation.Detail };
        }
        return Query(handle.DangerousGetHandle(), instanceId);
    }

    internal static InterruptModeResult Query(IntPtr handle, string instanceId)
    {
        if (QaRun.TryBlock("PCI INTERRUPT QUERY")) return InterruptModeResult.Pending;
        if (!instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
            return new(InterruptMode.NotPci, "This device has no PCI interrupt mode.");
        if (!NativeCfgMgr32.TryGetPciPdoName(instanceId, out string name, out string error))
            return new(InterruptMode.Unavailable, error, DateTimeOffset.Now);
        if (!TryBuildRequest(name, instanceId, out byte[] request))
            return new(InterruptMode.Unavailable, "Invalid PDO name / PCI instance ID", DateTimeOffset.Now);
        byte[] reply = new byte[PciInterruptMode.ReplySize];
        if (!DeviceIoControl(handle, IoctlQueryPci, request, request.Length, reply, reply.Length, out int returned, IntPtr.Zero))
        {
            int code = Marshal.GetLastWin32Error();
            string detail = code is 1 or 50 ? "Loaded DTIMOD does not support PCI queries. Restart Windows, then read IRQ Mode again."
                : $"DTIMOD PCI query: {new Win32Exception(code).Message} ({code})";
            return new(InterruptMode.Unavailable, detail, DateTimeOffset.Now);
        }
        if (returned < 0 || returned > reply.Length) return new(InterruptMode.Unavailable, "Invalid DTIMOD reply size", DateTimeOffset.Now);
        return PciInterruptMode.DecodeReply(instanceId, reply.AsSpan(0, returned));
    }

    internal static bool TryBuildRequest(string pdoName, string instanceId, out byte[] request)
    {
        request = [];
        if (!pdoName.StartsWith(@"\Device\", StringComparison.OrdinalIgnoreCase)
            || !instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)
            || pdoName.Length is 0 or >= 256 || instanceId.Length is 0 or >= 256
            || pdoName.Contains('\0') || instanceId.Contains('\0')) return false;
        request = new byte[1028];
        request[0] = 1;
        Encoding.Unicode.GetBytes(pdoName).CopyTo(request, 4);
        Encoding.Unicode.GetBytes(instanceId).CopyTo(request, 516);
        return true;
    }
}
