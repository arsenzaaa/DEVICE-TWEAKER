using System.Buffers.Binary;
using System.Text;

namespace DeviceTweakerCS;

// IOCTL_IMOD_WRITE_DEVICE request from IMOD\Driver\imod_driver.h. DTIMOD writes only
// when the address lies inside a memory range Windows assigned to this exact devnode.
internal static class ImodDeviceWrite
{
    internal const uint Version = 1;
    internal const uint IoctlCode = (0x8010u << 16) | (2u << 14) | ((0x810u + 5u) << 2);
    internal const int RequestSize = 1048;
    // ULONG version + ULONG resourcesBytes + UCHAR resources[65536].
    internal const int ReplySize = 8 + 65536;
    private const int PdoNameOffset = 24;
    private const int InstanceIdOffset = 536;
    private const int NameChars = 256;

    // Windows error codes DeviceIoControl reports for the driver's statuses.
    internal const int ErrorAccessDenied = 5;
    internal const int ErrorInvalidFunction = 1;

    internal static bool TryBuild(string pdoName, string instanceId, ulong address, uint size, ulong value, out byte[] request)
    {
        request = [];
        if (!pdoName.StartsWith(@"\Device\", StringComparison.OrdinalIgnoreCase)
            || !instanceId.StartsWith(@"PCI\", StringComparison.OrdinalIgnoreCase)
            || pdoName.Length >= NameChars || instanceId.Length >= NameChars
            || pdoName.Contains('\0') || instanceId.Contains('\0')
            || size is not (1 or 2 or 4))
        {
            return false;
        }

        request = new byte[RequestSize];
        BinaryPrimitives.WriteUInt32LittleEndian(request.AsSpan(0), Version);
        BinaryPrimitives.WriteUInt32LittleEndian(request.AsSpan(4), size);
        BinaryPrimitives.WriteUInt64LittleEndian(request.AsSpan(8), address);
        BinaryPrimitives.WriteUInt64LittleEndian(request.AsSpan(16), value);
        Encoding.Unicode.GetBytes(pdoName).CopyTo(request, PdoNameOffset);
        Encoding.Unicode.GetBytes(instanceId).CopyTo(request, InstanceIdOffset);
        return true;
    }

    internal static string DescribeFailure(int win32Error, string fallback, ulong address = 0, string? instanceId = null) => win32Error switch
    {
        ErrorAccessDenied => $"DTIMOD refused the write: 0x{address:X} is outside the memory of {instanceId ?? "the device"}",
        ErrorInvalidFunction => DriverRuntime.RestartMessage,
        _ => fallback,
    };
}
