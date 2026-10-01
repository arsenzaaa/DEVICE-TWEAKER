using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace DeviceTweakerCS;

internal static class UsbTopologyInterop
{
    private const uint FileDeviceUsb = 0x22;
    private const uint FileAnyAccess = 0;
    private const uint MethodBuffered = 0;

    private const uint DigcfPresent = 0x00000002;
    private const uint DigcfDeviceInterface = 0x00000010;

    private const uint GenericRead = 0x80000000;
    private const uint GenericWrite = 0x40000000;
    private const uint FileShareRead = 0x00000001;
    private const uint FileShareWrite = 0x00000002;
    private const uint OpenExisting = 3;

    private const int ErrorInsufficientBuffer = 122;
    private const int ErrorNoMoreItems = 259;

    private const int UsbConfigurationDescriptorType = 2;
    private const int UsbInterfaceDescriptorType = 4;
    private const int UsbEndpointDescriptorType = 5;

    private const int UsbNodeInformationBufferSize = 128;
    private const int UsbNodeConnectionInformationExSize = 35;
    private const int UsbDescriptorRequestHeaderSize = 12;
    private const int LargeNameBufferSize = 4096;
    private const int MaxTopologyDepth = 32;

    private static readonly Guid GuidDevinterfaceUsbHostController = new("3ABF6F2D-71C4-462A-8A92-1E6861E6AF27");
    private static readonly IntPtr InvalidHandleValue = new(-1);

    private static readonly uint IoctlUsbGetRootHubName =
        CtlCode(FileDeviceUsb, 258, MethodBuffered, FileAnyAccess);

    private static readonly uint IoctlUsbGetNodeInformation =
        CtlCode(FileDeviceUsb, 258, MethodBuffered, FileAnyAccess);

    private static readonly uint IoctlUsbGetDescriptorFromNodeConnection =
        CtlCode(FileDeviceUsb, 260, MethodBuffered, FileAnyAccess);

    private static readonly uint IoctlUsbGetNodeConnectionName =
        CtlCode(FileDeviceUsb, 261, MethodBuffered, FileAnyAccess);

    private static readonly uint IoctlUsbGetNodeConnectionDriverKeyName =
        CtlCode(FileDeviceUsb, 264, MethodBuffered, FileAnyAccess);

    private static readonly uint IoctlUsbGetNodeConnectionInformationEx =
        CtlCode(FileDeviceUsb, 274, MethodBuffered, FileAnyAccess);

    [StructLayout(LayoutKind.Sequential)]
    private struct SP_DEVICE_INTERFACE_DATA
    {
        public int cbSize;
        public Guid InterfaceClassGuid;
        public int Flags;
        public IntPtr Reserved;
    }

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern IntPtr SetupDiGetClassDevs(
        ref Guid classGuid,
        IntPtr enumerator,
        IntPtr hwndParent,
        uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    private static extern bool SetupDiEnumDeviceInterfaces(
        IntPtr deviceInfoSet,
        IntPtr deviceInfoData,
        ref Guid interfaceClassGuid,
        uint memberIndex,
        ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetupDiGetDeviceInterfaceDetail(
        IntPtr deviceInfoSet,
        ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
        IntPtr deviceInterfaceDetailData,
        uint deviceInterfaceDetailDataSize,
        out uint requiredSize,
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

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(
        SafeFileHandle device,
        uint ioControlCode,
        IntPtr inBuffer,
        int inBufferSize,
        [In, Out] byte[] outBuffer,
        int outBufferSize,
        out int bytesReturned,
        IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(
        SafeFileHandle device,
        uint ioControlCode,
        [In, Out] byte[] inBuffer,
        int inBufferSize,
        [In, Out] byte[] outBuffer,
        int outBufferSize,
        out int bytesReturned,
        IntPtr overlapped);

    public static List<UsbEndpointInfo> EnumerateEndpoints()
    {
        List<UsbEndpointInfo> endpoints = [];
        List<string> hostControllers = EnumerateHostControllerPaths();

        for (int i = 0; i < hostControllers.Count; i++)
        {
            string hostControllerPath = hostControllers[i];
            try
            {
                using SafeFileHandle hostController = OpenDevicePath(hostControllerPath);
                if (hostController.IsInvalid)
                {
                    continue;
                }

                string? rootHubName = QueryRootHubName(hostController);
                if (string.IsNullOrWhiteSpace(rootHubName))
                {
                    continue;
                }

                using SafeFileHandle rootHub = OpenUsbSymbolicName(rootHubName);
                if (rootHub.IsInvalid)
                {
                    continue;
                }

                string rootHubPath = NormalizeUsbSymbolicName(rootHubName);
                EnumerateHub(rootHub, rootHubPath, hostControllerPath, $"HC{i}", endpoints, 0);
            }
            catch
            {
            }
        }

        return endpoints;
    }

    private static List<string> EnumerateHostControllerPaths()
    {
        List<string> paths = [];
        Guid guid = GuidDevinterfaceUsbHostController;
        IntPtr infoSet = SetupDiGetClassDevs(ref guid, IntPtr.Zero, IntPtr.Zero, DigcfPresent | DigcfDeviceInterface);
        if (infoSet == InvalidHandleValue || infoSet == IntPtr.Zero)
        {
            return paths;
        }

        try
        {
            for (uint index = 0; ; index++)
            {
                SP_DEVICE_INTERFACE_DATA data = new()
                {
                    cbSize = Marshal.SizeOf<SP_DEVICE_INTERFACE_DATA>(),
                };

                if (!SetupDiEnumDeviceInterfaces(infoSet, IntPtr.Zero, ref guid, index, ref data))
                {
                    if (Marshal.GetLastWin32Error() == ErrorNoMoreItems)
                    {
                        break;
                    }

                    break;
                }

                _ = SetupDiGetDeviceInterfaceDetail(infoSet, ref data, IntPtr.Zero, 0, out uint requiredSize, IntPtr.Zero);
                if (requiredSize == 0 || Marshal.GetLastWin32Error() != ErrorInsufficientBuffer)
                {
                    continue;
                }

                IntPtr detailBuffer = Marshal.AllocHGlobal((int)requiredSize);
                try
                {
                    Marshal.WriteInt32(detailBuffer, IntPtr.Size == 8 ? 8 : 6);
                    if (!SetupDiGetDeviceInterfaceDetail(infoSet, ref data, detailBuffer, requiredSize, out _, IntPtr.Zero))
                    {
                        continue;
                    }

                    string? path = Marshal.PtrToStringUni(IntPtr.Add(detailBuffer, 4));
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        paths.Add(path);
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
            _ = SetupDiDestroyDeviceInfoList(infoSet);
        }

        return paths;
    }

    private static void EnumerateHub(
        SafeFileHandle hub,
        string hubPath,
        string hostControllerPath,
        string topologyPrefix,
        List<UsbEndpointInfo> endpoints,
        int depth)
    {
        if (depth > MaxTopologyDepth)
        {
            return;
        }

        int portCount = QueryHubPortCount(hub);
        if (portCount <= 0)
        {
            return;
        }

        for (int port = 1; port <= portCount; port++)
        {
            string scope = $"{topologyPrefix}/Port{port}";
            byte[]? connection = QueryConnectionInfo(hub, port);
            if (connection is null)
            {
                continue;
            }

            int connectionStatus = ToInt32(connection, 31);
            if (connectionStatus != 1)
            {
                continue;
            }

            bool isHub = connection[24] != 0;
            int deviceAddress = ToUInt16(connection, 25);
            int currentConfigurationValue = connection[22];
            int maxPacket0 = connection[11];
            string speed = UsbSpeedToString(connection[23]);

            byte[] deviceDescriptor = new byte[18];
            Buffer.BlockCopy(connection, 4, deviceDescriptor, 0, deviceDescriptor.Length);

            string vendorId = ToUInt16(deviceDescriptor, 8).ToString("X4");
            string productId = ToUInt16(deviceDescriptor, 10).ToString("X4");
            int configurationCount = deviceDescriptor[17];

            if (configurationCount > 0)
            {
                byte[]? config = FindActiveConfigurationDescriptor(hub, port, (byte)currentConfigurationValue, (byte)configurationCount);
                if (config is not null)
                {
                    ParseConfigurationEndpoints(
                        config,
                        hostControllerPath,
                        hubPath,
                        scope,
                        port,
                        speed,
                        isHub,
                        deviceAddress,
                        vendorId,
                        productId,
                        currentConfigurationValue,
                        maxPacket0,
                        endpoints);
                }
            }

            if (!isHub)
            {
                continue;
            }

            string? childHubName = QueryConnectionUnicodeField(hub, IoctlUsbGetNodeConnectionName, port, 4);
            if (string.IsNullOrWhiteSpace(childHubName))
            {
                continue;
            }

            try
            {
                using SafeFileHandle childHub = OpenUsbSymbolicName(childHubName);
                if (!childHub.IsInvalid)
                {
                    EnumerateHub(childHub, NormalizeUsbSymbolicName(childHubName), hostControllerPath, scope, endpoints, depth + 1);
                }
            }
            catch
            {
            }
        }
    }

    private static void ParseConfigurationEndpoints(
        byte[] config,
        string hostControllerPath,
        string hubPath,
        string topologyPath,
        int portNumber,
        string speed,
        bool deviceIsHub,
        int deviceAddress,
        string vendorId,
        string productId,
        int currentConfigurationValue,
        int maxPacket0,
        List<UsbEndpointInfo> endpoints)
    {
        _ = currentConfigurationValue;
        _ = maxPacket0;

        int interfaceNumber = -1;
        int alternateSetting = -1;
        string interfaceClass = string.Empty;
        string interfaceSubClass = string.Empty;
        string interfaceProtocol = string.Empty;

        for (int offset = 0; offset + 2 <= config.Length;)
        {
            int length = config[offset];
            int descriptorType = config[offset + 1];

            if (length <= 0 || offset + length > config.Length)
            {
                break;
            }

            if (descriptorType == UsbInterfaceDescriptorType && length >= 9)
            {
                interfaceNumber = config[offset + 2];
                alternateSetting = config[offset + 3];
                interfaceClass = "0x" + config[offset + 5].ToString("X2");
                interfaceSubClass = "0x" + config[offset + 6].ToString("X2");
                interfaceProtocol = "0x" + config[offset + 7].ToString("X2");
            }
            else if (descriptorType == UsbEndpointDescriptorType && length >= 7)
            {
                byte endpointAddress = config[offset + 2];
                byte attributes = config[offset + 3];

                endpoints.Add(new UsbEndpointInfo
                {
                    HostControllerPath = hostControllerPath,
                    HubPath = hubPath,
                    TopologyPath = topologyPath,
                    PortNumber = portNumber,
                    Speed = speed,
                    DeviceIsHub = deviceIsHub,
                    DeviceAddress = deviceAddress,
                    VendorId = vendorId,
                    ProductId = productId,
                    InterfaceNumber = interfaceNumber,
                    AlternateSetting = alternateSetting,
                    InterfaceClass = interfaceClass,
                    InterfaceSubClass = interfaceSubClass,
                    InterfaceProtocol = interfaceProtocol,
                    Direction = (endpointAddress & 0x80) != 0 ? "IN" : "OUT",
                    TransferType = TransferTypeToString(attributes & 0x03),
                    BInterval = config[offset + 6],
                });
            }

            offset += length;
        }
    }

    private static byte[]? FindActiveConfigurationDescriptor(SafeFileHandle hub, int port, byte currentConfigurationValue, byte configurationCount)
    {
        if (currentConfigurationValue == 0)
        {
            return null;
        }

        for (byte index = 0; index < configurationCount; index++)
        {
            byte[]? config = QueryConfigurationDescriptor(hub, port, index);
            if (config is null || config.Length < 9)
            {
                continue;
            }

            if (config[5] == currentConfigurationValue)
            {
                return config;
            }
        }

        return null;
    }

    private static byte[]? QueryConfigurationDescriptor(SafeFileHandle hub, int port, byte descriptorIndex)
    {
        byte[] first = BuildDescriptorRequest(port, UsbConfigurationDescriptorType, descriptorIndex, 9);
        if (!DeviceIoControl(hub, IoctlUsbGetDescriptorFromNodeConnection, first, first.Length, first, first.Length, out int bytesReturned, IntPtr.Zero)
            || bytesReturned != first.Length)
        {
            return null;
        }

        int totalLength = ToUInt16(first, UsbDescriptorRequestHeaderSize + 2);
        if (totalLength < 9)
        {
            return null;
        }

        byte[] full = BuildDescriptorRequest(port, UsbConfigurationDescriptorType, descriptorIndex, totalLength);
        if (!DeviceIoControl(hub, IoctlUsbGetDescriptorFromNodeConnection, full, full.Length, full, full.Length, out bytesReturned, IntPtr.Zero)
            || bytesReturned != full.Length)
        {
            return null;
        }

        byte[] config = new byte[totalLength];
        Buffer.BlockCopy(full, UsbDescriptorRequestHeaderSize, config, 0, totalLength);
        return config;
    }

    private static byte[] BuildDescriptorRequest(int port, int descriptorType, byte descriptorIndex, int dataLength)
    {
        byte[] buffer = new byte[UsbDescriptorRequestHeaderSize + dataLength];
        WriteUInt32(buffer, 0, (uint)port);
        WriteUInt16(buffer, 6, (ushort)(((descriptorType & 0xFF) << 8) | descriptorIndex));
        WriteUInt16(buffer, 10, (ushort)dataLength);
        return buffer;
    }

    private static int QueryHubPortCount(SafeFileHandle hub)
    {
        byte[] buffer = new byte[UsbNodeInformationBufferSize];
        int bytesReturned;
        return DeviceIoControl(hub, IoctlUsbGetNodeInformation, buffer, buffer.Length, buffer, buffer.Length, out bytesReturned, IntPtr.Zero)
               && bytesReturned >= 7
            ? buffer[6]
            : -1;
    }

    private static byte[]? QueryConnectionInfo(SafeFileHandle hub, int port)
    {
        byte[] buffer = new byte[UsbNodeConnectionInformationExSize];
        WriteUInt32(buffer, 0, (uint)port);
        return DeviceIoControl(hub, IoctlUsbGetNodeConnectionInformationEx, buffer, buffer.Length, buffer, buffer.Length, out int bytesReturned, IntPtr.Zero)
               && bytesReturned >= UsbNodeConnectionInformationExSize
            ? buffer
            : null;
    }

    private static string? QueryRootHubName(SafeFileHandle hostController)
    {
        byte[] buffer = new byte[LargeNameBufferSize];
        if (!DeviceIoControl(hostController, IoctlUsbGetRootHubName, IntPtr.Zero, 0, buffer, buffer.Length, out int bytesReturned, IntPtr.Zero)
            || bytesReturned < 4)
        {
            return null;
        }

        int actualLength = ToInt32(buffer, 0);
        if (actualLength < 4 || actualLength > buffer.Length)
        {
            return null;
        }

        return DecodeUnicodeString(buffer, 4, actualLength - 4);
    }

    private static string? QueryConnectionUnicodeField(SafeFileHandle hub, uint ioctl, int port, int lengthOffset)
    {
        byte[] buffer = new byte[LargeNameBufferSize];
        WriteUInt32(buffer, 0, (uint)port);
        if (!DeviceIoControl(hub, ioctl, buffer, buffer.Length, buffer, buffer.Length, out int bytesReturned, IntPtr.Zero)
            || bytesReturned < lengthOffset + 4)
        {
            return null;
        }

        int actualLength = ToInt32(buffer, lengthOffset);
        int stringOffset = lengthOffset + 4;
        if (actualLength <= stringOffset || actualLength > buffer.Length)
        {
            return null;
        }

        return DecodeUnicodeString(buffer, stringOffset, actualLength - stringOffset);
    }

    private static SafeFileHandle OpenDevicePath(string path)
    {
        return CreateFile(
            path,
            GenericRead | GenericWrite,
            FileShareRead | FileShareWrite,
            IntPtr.Zero,
            OpenExisting,
            0,
            IntPtr.Zero);
    }

    private static SafeFileHandle OpenUsbSymbolicName(string symbolicName)
    {
        return OpenDevicePath(NormalizeUsbSymbolicName(symbolicName));
    }

    private static string NormalizeUsbSymbolicName(string symbolicName)
    {
        string path = symbolicName.Trim();
        if (path.StartsWith(@"\??\", StringComparison.Ordinal))
        {
            path = path[4..];
        }

        if (path.StartsWith(@"\\.\", StringComparison.Ordinal)
            || path.StartsWith(@"\\?\", StringComparison.Ordinal))
        {
            return path;
        }

        return @"\\.\" + path.TrimStart('\\');
    }

    private static uint CtlCode(uint deviceType, uint function, uint method, uint access)
    {
        return (deviceType << 16) | (access << 14) | (function << 2) | method;
    }

    private static int ToUInt16(byte[] buffer, int offset)
    {
        return offset + 1 < buffer.Length
            ? buffer[offset] | (buffer[offset + 1] << 8)
            : 0;
    }

    private static int ToInt32(byte[] buffer, int offset)
    {
        return offset + 3 < buffer.Length
            ? buffer[offset] | (buffer[offset + 1] << 8) | (buffer[offset + 2] << 16) | (buffer[offset + 3] << 24)
            : 0;
    }

    private static void WriteUInt16(byte[] buffer, int offset, ushort value)
    {
        buffer[offset] = (byte)(value & 0xFF);
        buffer[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] buffer, int offset, uint value)
    {
        buffer[offset] = (byte)(value & 0xFF);
        buffer[offset + 1] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static string DecodeUnicodeString(byte[] buffer, int offset, int byteLength)
    {
        if (byteLength <= 0 || offset < 0 || offset + byteLength > buffer.Length)
        {
            return string.Empty;
        }

        return System.Text.Encoding.Unicode.GetString(buffer, offset, byteLength).TrimEnd('\0');
    }

    private static string UsbSpeedToString(int speed)
    {
        return speed switch
        {
            0 => "Low",
            1 => "Full",
            2 => "High",
            3 => "Super",
            _ => "Unknown",
        };
    }

    private static string TransferTypeToString(int transferType)
    {
        return transferType switch
        {
            0 => "Control",
            1 => "Isochronous",
            2 => "Bulk",
            3 => "Interrupt",
            _ => "Unknown",
        };
    }
}
