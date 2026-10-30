using System.Runtime.InteropServices;
using System.Text;

namespace DeviceTweakerCS;

internal static class NativeCfgMgr32
{
    private const int CR_SUCCESS = 0x00000000;
    private static readonly Guid PciDevicePropertyGuid = new("3AB22E31-8264-4B4E-9AF5-A8D2D8E33E62");

    [StructLayout(LayoutKind.Sequential)]
    private struct DevPropKey
    {
        public Guid FormatId;
        public uint PropertyId;
    }

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int CM_Locate_DevNodeW(out uint pdnDevInst, string pDeviceID, uint ulFlags);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    private static extern int CM_Get_Parent(out uint pdnDevInst, uint dnDevInst, uint ulFlags);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    private static extern int CM_Get_Device_ID_Size(out uint pulLen, uint dnDevInst, uint ulFlags);

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int CM_Get_Device_IDW(uint dnDevInst, StringBuilder buffer, int bufferLen, uint ulFlags);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    private static extern int CM_Get_DevNode_PropertyW(
        uint dnDevInst,
        ref DevPropKey propertyKey,
        out uint propertyType,
        byte[] propertyBuffer,
        ref uint propertyBufferSize,
        uint flags);

    public static bool TryGetPciInterruptCapabilities(string instanceId, out PciInterruptCapabilities? capabilities)
    {
        capabilities = null;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        try
        {
            // A bridge's capabilities do not describe its child device.
            if (!instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)
                || CM_Locate_DevNodeW(out uint devInst, instanceId, 0) != CR_SUCCESS
                || !TryGetDevNodeUInt32(devInst, 14, out uint supportMask)) return false;
            uint? maximum = TryGetDevNodeUInt32(devInst, 15, out uint maxValue) ? maxValue : null;
            capabilities = new PciInterruptCapabilities((PciInterruptSupport)(supportMask & 0x7), maximum);
            return true;
        }
        catch (Exception ex)
        {
            AppDiagnostics.Warn("PCI.CAPS", ex);
        }

        return false;
    }

    internal static bool TryGetPciPdoName(string instanceId, out string pdoName, out string error)
    {
        pdoName = string.Empty;
        error = string.Empty;
        if (!instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
        {
            error = "not a PCI device";
            return false;
        }
        int cr = CM_Locate_DevNodeW(out uint devInst, instanceId, 0);
        if (cr != CR_SUCCESS) { error = $"CM_Locate_DevNode: 0x{cr:X}"; return false; }
        // DEVPKEY_Device_PDOName; exact devnode, no parent traversal.
        DevPropKey key = new() { FormatId = new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), PropertyId = 16 };
        byte[] data = new byte[512];
        uint size = (uint)data.Length;
        cr = CM_Get_DevNode_PropertyW(devInst, ref key, out uint type, data, ref size, 0);
        if (cr != CR_SUCCESS || type != 0x12 || size < 2 || size > data.Length || size % 2 != 0)
        { error = $"CM_Get_DevNode_Property(PDOName): 0x{cr:X}, type={type}, bytes={size}"; return false; }
        if (data[size - 2] != 0 || data[size - 1] != 0)
        { error = "PDO name is not terminated"; return false; }
        pdoName = Encoding.Unicode.GetString(data, 0, (int)size - 2);
        if (string.IsNullOrEmpty(pdoName)) { error = "Windows returned an empty PDO name"; return false; }
        return true;
    }

    private static bool TryGetDevNodeUInt32(uint devInst, uint propertyId, out uint value)
    {
        value = 0;
        DevPropKey key = new() { FormatId = PciDevicePropertyGuid, PropertyId = propertyId };
        byte[] data = new byte[sizeof(uint)];
        uint size = (uint)data.Length;
        int cr = CM_Get_DevNode_PropertyW(devInst, ref key, out uint propertyType, data, ref size, 0);
        // DEVPROP_TYPE_UINT32 = 0x00000007.
        if (cr != CR_SUCCESS || propertyType != 0x00000007 || size < sizeof(uint))
        {
            return false;
        }

        value = BitConverter.ToUInt32(data, 0);
        return true;
    }

    public static bool TryGetParentInstanceId(string instanceId, out string? parentInstanceId)
    {
        parentInstanceId = null;
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return false;
        }

        try
        {
            int cr = CM_Locate_DevNodeW(out uint devInst, instanceId, 0);
            if (cr != CR_SUCCESS)
            {
                return false;
            }

            cr = CM_Get_Parent(out uint parentDevInst, devInst, 0);
            if (cr != CR_SUCCESS)
            {
                return false;
            }

            cr = CM_Get_Device_ID_Size(out uint idLen, parentDevInst, 0);
            if (cr != CR_SUCCESS)
            {
                return false;
            }

            StringBuilder buffer = new((int)idLen + 1);
            cr = CM_Get_Device_IDW(parentDevInst, buffer, buffer.Capacity, 0);
            if (cr != CR_SUCCESS)
            {
                return false;
            }

            parentInstanceId = buffer.ToString();
            return !string.IsNullOrWhiteSpace(parentInstanceId);
        }
        catch
        {
            return false;
        }
    }
}

