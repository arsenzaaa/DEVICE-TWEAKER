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
            string currentId = instanceId;
            for (int depth = 0; depth < 8; depth++)
            {
                if (CM_Locate_DevNodeW(out uint devInst, currentId, 0) != CR_SUCCESS)
                {
                    return false;
                }

                if (TryGetDevNodeUInt32(devInst, 14, out uint supportMask))
                {
                    uint? maximum = TryGetDevNodeUInt32(devInst, 15, out uint maxValue)
                        ? maxValue
                        : null;
                    capabilities = new PciInterruptCapabilities(
                        (PciInterruptSupport)(supportMask & 0x7),
                        maximum);
                    return true;
                }

                if (CM_Get_Parent(out uint parentDevInst, devInst, 0) != CR_SUCCESS
                    || CM_Get_Device_ID_Size(out uint idLen, parentDevInst, 0) != CR_SUCCESS)
                {
                    return false;
                }

                StringBuilder parentId = new((int)idLen + 1);
                if (CM_Get_Device_IDW(parentDevInst, parentId, parentId.Capacity, 0) != CR_SUCCESS)
                {
                    return false;
                }

                currentId = parentId.ToString();
            }
        }
        catch
        {
        }

        return false;
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

