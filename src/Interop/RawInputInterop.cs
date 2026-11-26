using System.Collections.Concurrent;
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

    private const uint RidHeader = 0x10000005;
    private const uint RidiDeviceName = 0x20000007;
    private const uint RidevRemove = 0x00000001;
    private const uint RidevInputSink = 0x00000100;

    private const ushort UsagePageGeneric = 0x01;
    private const ushort UsageMouse = 0x02;
    private const ushort UsageKeyboard = 0x06;

    private const uint RimTypeMouse = 0;
    private const uint RimTypeKeyboard = 1;

    private static readonly uint HeaderSize = (uint)Marshal.SizeOf<RAWINPUTHEADER>();
    private static readonly uint DeviceStructSize = (uint)Marshal.SizeOf<RAWINPUTDEVICE>();
    private static readonly ConcurrentDictionary<IntPtr, (string DeviceName, string InstanceId)> DeviceNameCache = new();

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
        out RAWINPUTHEADER pData,
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

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMsg
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public Point pt;
    }

    [DllImport("user32.dll")]
    private static extern int GetMessageW(out NativeMsg lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessageW(ref NativeMsg lpMsg);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessageW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern void PostQuitMessage(int nExitCode);

    internal sealed class RawInputBackgroundMonitor : IDisposable
    {
        private const int WmClose = 0x0010;
        private const int WmStopMonitor = 0x8001;
        private static readonly IntPtr HwndMessage = new(-3);

        private readonly Action<RawInputDeviceKind, IntPtr> _onSample;
        private readonly ManualResetEventSlim _started = new(false);
        private readonly Thread _thread;
        private IntPtr _hwnd;
        private bool _registered;
        private volatile bool _stopped;

        internal bool Registered => _registered;

        internal RawInputBackgroundMonitor(Action<RawInputDeviceKind, IntPtr> onSample)
        {
            _onSample = onSample;
            _thread = new Thread(RunLoop)
            {
                IsBackground = true,
                Name = "DEVICE TWEAKER raw input",
            };
            _thread.Start();
            _started.Wait();
            _started.Dispose();
        }

        private void RunLoop()
        {
            MessageOnlyWindow? window = null;
            bool signaled = false;
            try
            {
                window = new MessageOnlyWindow(this);
                Volatile.Write(ref _hwnd, window.Handle);
                if (window.Handle == IntPtr.Zero || !RegisterMouseAndKeyboard(window.Handle))
                {
                    signaled = true;
                    _started.Set();
                    return;
                }

                _registered = true;
                signaled = true;
                _started.Set();

                if (_stopped)
                {
                    return;
                }

                while (GetMessageW(out NativeMsg msg, IntPtr.Zero, 0, 0) > 0)
                {
                    _ = DispatchMessageW(ref msg);
                }
            }
            catch
            {
                if (!signaled)
                {
                    signaled = true;
                    _started.Set();
                }
            }
            finally
            {
                Volatile.Write(ref _hwnd, IntPtr.Zero);
                if (_registered)
                {
                    _ = UnregisterMouseAndKeyboardSafe();
                }

                window?.Destroy();
            }
        }

        private static bool UnregisterMouseAndKeyboardSafe()
        {
            try
            {
                UnregisterMouseAndKeyboard();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            if (_stopped)
            {
                return;
            }

            _stopped = true;
            IntPtr hwnd = Volatile.Read(ref _hwnd);
            if (hwnd != IntPtr.Zero)
            {
                _ = PostMessageW(hwnd, WmStopMonitor, IntPtr.Zero, IntPtr.Zero);
            }

            if (Thread.CurrentThread != _thread && _thread.IsAlive)
            {
                _ = _thread.Join(50);
            }
        }

        private sealed class MessageOnlyWindow : NativeWindow
        {
            private readonly RawInputBackgroundMonitor _owner;

            internal MessageOnlyWindow(RawInputBackgroundMonitor owner)
            {
                _owner = owner;
                CreateHandle(new CreateParams
                {
                    Caption = string.Empty,
                    Parent = HwndMessage,
                });
            }

            internal void Destroy()
            {
                if (Handle != IntPtr.Zero)
                {
                    DestroyHandle();
                }
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmInput && !_owner._stopped)
                {
                    if (TryReadHeader(m.LParam, out RawInputDeviceKind kind, out IntPtr deviceHandle))
                    {
                        try
                        {
                            _owner._onSample(kind, deviceHandle);
                        }
                        catch
                        {
                            // Keep the raw input pump alive if a single sample callback fails during shutdown.
                        }
                    }

                    base.WndProc(ref m);
                    return;
                }

                if (m.Msg is WmStopMonitor or WmClose)
                {
                    PostQuitMessage(0);
                    m.Result = IntPtr.Zero;
                    return;
                }

                base.WndProc(ref m);
            }
        }
    }

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

        return RegisterRawInputDevices(devices, (uint)devices.Length, DeviceStructSize);
    }

    public static void UnregisterMouseAndKeyboard()
    {
        RAWINPUTDEVICE[] devices =
        [
            new()
            {
                usUsagePage = UsagePageGeneric,
                usUsage = UsageMouse,
                dwFlags = RidevRemove,
                hwndTarget = IntPtr.Zero,
            },
            new()
            {
                usUsagePage = UsagePageGeneric,
                usUsage = UsageKeyboard,
                dwFlags = RidevRemove,
                hwndTarget = IntPtr.Zero,
            },
        ];

        _ = RegisterRawInputDevices(devices, (uint)devices.Length, DeviceStructSize);
    }

    // Zero-allocation hot path for WM_INPUT (1K-8K mice fire thousands of times/sec).
    public static bool TryReadHeader(IntPtr lParam, out RawInputDeviceKind kind, out IntPtr deviceHandle)
    {
        uint size = HeaderSize;
        uint result = GetRawInputData(lParam, RidHeader, out RAWINPUTHEADER header, ref size, HeaderSize);
        if (result == uint.MaxValue || result < HeaderSize || header.hDevice == IntPtr.Zero)
        {
            kind = RawInputDeviceKind.Other;
            deviceHandle = IntPtr.Zero;
            return false;
        }

        kind = header.dwType switch
        {
            RimTypeMouse => RawInputDeviceKind.Mouse,
            RimTypeKeyboard => RawInputDeviceKind.Keyboard,
            _ => RawInputDeviceKind.Other,
        };
        deviceHandle = header.hDevice;
        return kind != RawInputDeviceKind.Other;
    }

    public static bool TryGetDeviceIdentity(IntPtr deviceHandle, out string deviceName, out string instanceId)
    {
        (deviceName, instanceId) = GetDeviceIdentity(deviceHandle);
        return !string.IsNullOrWhiteSpace(instanceId);
    }

    public static bool TryGetMessage(IntPtr lParam, out RawInputMessage message)
    {
        if (!TryReadHeader(lParam, out RawInputDeviceKind kind, out IntPtr deviceHandle)
            || !TryGetDeviceIdentity(deviceHandle, out string deviceName, out string instanceId))
        {
            message = new RawInputMessage(RawInputDeviceKind.Other, IntPtr.Zero, string.Empty, string.Empty);
            return false;
        }

        message = new RawInputMessage(kind, deviceHandle, deviceName, instanceId);
        return true;
    }

    private static (string DeviceName, string InstanceId) GetDeviceIdentity(IntPtr deviceHandle)
    {
        return DeviceNameCache.GetOrAdd(deviceHandle, static handle =>
        {
            string deviceName = GetDeviceName(handle);
            string instanceId = TryParseInstanceIdFromRawDeviceName(deviceName) ?? string.Empty;
            return (deviceName, instanceId);
        });
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
