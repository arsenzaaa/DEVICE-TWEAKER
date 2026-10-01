using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32.SafeHandles;

namespace DeviceTweakerCS;

internal static partial class HidInterop
{
    private const uint DIGCF_PRESENT = 0x00000002;
    private const uint DIGCF_DEVICEINTERFACE = 0x00000010;
    private const int ERROR_NO_MORE_ITEMS = 259;

    private const uint FILE_SHARE_READ = 0x00000001;
    private const uint FILE_SHARE_WRITE = 0x00000002;
    private const uint OPEN_EXISTING = 3;

    private const int HIDP_STATUS_SUCCESS = 0x00110000;

    [StructLayout(LayoutKind.Sequential)]
    private struct SP_DEVICE_INTERFACE_DATA
    {
        public int cbSize;
        public Guid InterfaceClassGuid;
        public int Flags;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HIDP_CAPS
    {
        public ushort Usage;
        public ushort UsagePage;
        public ushort InputReportByteLength;
        public ushort OutputReportByteLength;
        public ushort FeatureReportByteLength;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
        public ushort[] Reserved;

        public ushort NumberLinkCollectionNodes;
        public ushort NumberInputButtonCaps;
        public ushort NumberInputValueCaps;
        public ushort NumberInputDataIndices;
        public ushort NumberOutputButtonCaps;
        public ushort NumberOutputValueCaps;
        public ushort NumberOutputDataIndices;
        public ushort NumberFeatureButtonCaps;
        public ushort NumberFeatureValueCaps;
        public ushort NumberFeatureDataIndices;
    }

    [DllImport("hid.dll", SetLastError = true)]
    private static extern void HidD_GetHidGuid(out Guid hidGuid);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern IntPtr SetupDiGetClassDevs(ref Guid classGuid, IntPtr enumerator, IntPtr hwndParent, uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern bool SetupDiEnumDeviceInterfaces(
        IntPtr deviceInfoSet,
        IntPtr deviceInfoData,
        ref Guid interfaceClassGuid,
        int memberIndex,
        ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetupDiGetDeviceInterfaceDetail(
        IntPtr deviceInfoSet,
        ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
        IntPtr deviceInterfaceDetailData,
        int deviceInterfaceDetailDataSize,
        ref int requiredSize,
        IntPtr deviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(
        string fileName,
        uint desiredAccess,
        uint shareMode,
        IntPtr securityAttributes,
        uint creationDisposition,
        uint flagsAndAttributes,
        IntPtr templateFile);

    [DllImport("hid.dll", SetLastError = true)]
    private static extern bool HidD_GetProductString(SafeFileHandle hidDeviceObject, byte[] buffer, int bufferLength);

    [DllImport("hid.dll", SetLastError = true)]
    private static extern bool HidD_GetPreparsedData(SafeFileHandle hidDeviceObject, out IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    private static extern bool HidD_FreePreparsedData(IntPtr preparsedData);

    [DllImport("hid.dll", SetLastError = true)]
    private static extern int HidP_GetCaps(IntPtr preparsedData, out HIDP_CAPS caps);

    public static IEnumerable<string> EnumerateHidDevicePaths()
    {
        HidD_GetHidGuid(out Guid hidGuid);
        IntPtr deviceInfoSet = SetupDiGetClassDevs(ref hidGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
        if (deviceInfoSet == IntPtr.Zero || deviceInfoSet == new IntPtr(-1))
        {
            yield break;
        }

        try
        {
            SP_DEVICE_INTERFACE_DATA interfaceData = new()
            {
                cbSize = Marshal.SizeOf<SP_DEVICE_INTERFACE_DATA>(),
            };

            for (int index = 0; ; index++)
            {
                bool ok = SetupDiEnumDeviceInterfaces(deviceInfoSet, IntPtr.Zero, ref hidGuid, index, ref interfaceData);
                if (!ok)
                {
                    int err = Marshal.GetLastWin32Error();
                    if (err == ERROR_NO_MORE_ITEMS)
                    {
                        yield break;
                    }

                    yield break;
                }

                int requiredSize = 0;
                SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref interfaceData, IntPtr.Zero, 0, ref requiredSize, IntPtr.Zero);
                if (requiredSize <= 0)
                {
                    continue;
                }

                IntPtr detailBuffer = Marshal.AllocHGlobal(requiredSize);
                try
                {
                    Marshal.WriteInt32(detailBuffer, IntPtr.Size == 8 ? 8 : 6);
                    ok = SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref interfaceData, detailBuffer, requiredSize, ref requiredSize, IntPtr.Zero);
                    if (!ok)
                    {
                        continue;
                    }

                    IntPtr devicePathPtr = detailBuffer + 4;
                    string? devicePath = Marshal.PtrToStringUni(devicePathPtr);
                    if (!string.IsNullOrWhiteSpace(devicePath))
                    {
                        yield return devicePath;
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(detailBuffer);
                }
            }
        }
        finally
        {
            SetupDiDestroyDeviceInfoList(deviceInfoSet);
        }
    }

    public static bool TryReadProductAndUsage(string devicePath, out string product, out int? usagePage, out int? usageId)
    {
        product = "<none>";
        usagePage = null;
        usageId = null;

        if (string.IsNullOrWhiteSpace(devicePath))
        {
            return false;
        }

        try
        {
            using SafeFileHandle handle = CreateFile(
                devicePath,
                0,
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                0,
                IntPtr.Zero);

            if (handle.IsInvalid)
            {
                return false;
            }

            byte[] buf = new byte[256];
            if (HidD_GetProductString(handle, buf, buf.Length))
            {
                product = Encoding.Unicode.GetString(buf).TrimEnd('\0');
            }

            if (HidD_GetPreparsedData(handle, out IntPtr preparsedData))
            {
                try
                {
                    int status = HidP_GetCaps(preparsedData, out HIDP_CAPS caps);
                    if (status == HIDP_STATUS_SUCCESS)
                    {
                        usagePage = caps.UsagePage;
                        usageId = caps.Usage;
                    }
                }
                finally
                {
                    HidD_FreePreparsedData(preparsedData);
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public static string? TryParseInstanceIdFromDevicePath(string devicePath)
    {
        if (string.IsNullOrWhiteSpace(devicePath))
        {
            return null;
        }

        Match match = HidInstanceRegex().Match(devicePath);
        if (!match.Success)
        {
            return null;
        }

        string fragment = match.Groups[1].Value;
        string inst = "HID\\" + fragment.Replace('#', '\\');
        return inst;
    }

    [GeneratedRegex("(?i)hid#([^#]+)#", RegexOptions.CultureInvariant)]
    private static partial Regex HidInstanceRegex();
}

