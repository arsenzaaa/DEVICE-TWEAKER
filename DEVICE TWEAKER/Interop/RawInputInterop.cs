using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal enum RawInputDeviceKind
{
    Other,
    Mouse,
    Keyboard,
}

internal sealed record RawInputMessage(
    RawInputDeviceKind Kind,
    IntPtr DeviceHandle,
    string DeviceName,
    string InstanceId);

internal static partial class RawInputInterop
{
    public const int WmInput = 0x00FF;

    private const uint RidInput = 0x10000003;
    private const uint RidiDeviceName = 0x20000007;
    private const uint RidevInputSink = 0x00000100;

    private const ushort UsagePageGeneric = 0x01;
    private const ushort UsageMouse = 0x02;
    private const ushort UsageKeyboard = 0x06;

    private const uint RimTypeMouse = 0;
    private const uint RimTypeKeyboard = 1;

    private static readonly Dictionary<IntPtr, (string DeviceName, string InstanceId)> DeviceNameCache = [];

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTDEVICE
    {
        public ushort usUsagePage;
        public ushort usUsage;
        public uint dwFlags;
        public IntPtr hwndTarget;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTHEADER
    {
        public uint dwType;
        public uint dwSize;
        public IntPtr hDevice;
        public IntPtr wParam;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(
        RAWINPUTDEVICE[] pRawInputDevices,
        uint uiNumDevices,
        uint cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputData(
        IntPtr hRawInput,
        uint uiCommand,
        IntPtr pData,
        ref uint pcbSize,
        uint cbSizeHeader);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetRawInputDeviceInfoW(
        IntPtr hDevice,
        uint uiCommand,
        IntPtr pData,
        ref uint pcbSize);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetRawInputDeviceInfoW(
        IntPtr hDevice,
        uint uiCommand,
        StringBuilder pData,
        ref uint pcbSize);

    public static bool RegisterMouseAndKeyboard(IntPtr hwnd)
    {
        RAWINPUTDEVICE[] devices =
        [
            new()
            {
                usUsagePage = UsagePageGeneric,
                usUsage = UsageMouse,
                dwFlags = RidevInputSink,
                hwndTarget = hwnd,
            },
            new()
            {
                usUsagePage = UsagePageGeneric,
                usUsage = UsageKeyboard,
                dwFlags = RidevInputSink,
                hwndTarget = hwnd,
            },
        ];

        return RegisterRawInputDevices(devices, (uint)devices.Length, (uint)Marshal.SizeOf<RAWINPUTDEVICE>());
    }

    public static bool TryGetMessage(IntPtr lParam, out RawInputMessage message)
    {
        message = new RawInputMessage(RawInputDeviceKind.Other, IntPtr.Zero, string.Empty, string.Empty);

        uint size = 0;
        _ = GetRawInputData(lParam, RidInput, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());
        if (size < Marshal.SizeOf<RAWINPUTHEADER>())
        {
            return false;
        }

        IntPtr buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            uint result = GetRawInputData(lParam, RidInput, buffer, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());
            if (result == uint.MaxValue || result == 0)
            {
                return false;
            }

            RAWINPUTHEADER header = Marshal.PtrToStructure<RAWINPUTHEADER>(buffer);
            RawInputDeviceKind kind = header.dwType switch
            {
                RimTypeMouse => RawInputDeviceKind.Mouse,
                RimTypeKeyboard => RawInputDeviceKind.Keyboard,
                _ => RawInputDeviceKind.Other,
            };

            if (kind == RawInputDeviceKind.Other || header.hDevice == IntPtr.Zero)
            {
                return false;
            }

            (string deviceName, string instanceId) = GetDeviceIdentity(header.hDevice);
            if (string.IsNullOrWhiteSpace(instanceId))
            {
                return false;
            }

            message = new RawInputMessage(kind, header.hDevice, deviceName, instanceId);
            return true;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static (string DeviceName, string InstanceId) GetDeviceIdentity(IntPtr deviceHandle)
    {
        if (DeviceNameCache.TryGetValue(deviceHandle, out (string DeviceName, string InstanceId) cached))
        {
            return cached;
        }

        string deviceName = GetDeviceName(deviceHandle);
        string instanceId = TryParseInstanceIdFromRawDeviceName(deviceName) ?? string.Empty;
        (string DeviceName, string InstanceId) identity = (deviceName, instanceId);
        DeviceNameCache[deviceHandle] = identity;
        return identity;
    }

    private static string GetDeviceName(IntPtr deviceHandle)
    {
        uint chars = 0;
        _ = GetRawInputDeviceInfoW(deviceHandle, RidiDeviceName, IntPtr.Zero, ref chars);
        if (chars == 0)
        {
            return string.Empty;
        }

        StringBuilder builder = new((int)chars + 1);
        uint result = GetRawInputDeviceInfoW(deviceHandle, RidiDeviceName, builder, ref chars);
        return result == uint.MaxValue ? string.Empty : builder.ToString();
    }

    private static string? TryParseInstanceIdFromRawDeviceName(string deviceName)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
        {
            return null;
        }

        Match match = HidInstanceRegex().Match(deviceName);
        if (!match.Success)
        {
            return null;
        }

        return "HID\\"
            + match.Groups[1].Value.Replace('#', '\\').ToUpperInvariant()
            + "\\"
            + match.Groups[2].Value.Replace('#', '\\').ToUpperInvariant();
    }

    [GeneratedRegex("(?i)hid#([^#]+)#([^#]+)#", RegexOptions.CultureInvariant)]
    private static partial Regex HidInstanceRegex();
}
