using System.Runtime.InteropServices;
using System.Text;

namespace DeviceTweakerCS;

internal static class NativeCfgMgr32
{
    private const int CR_SUCCESS = 0x00000000;

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int CM_Locate_DevNodeW(out uint pdnDevInst, string pDeviceID, uint ulFlags);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    private static extern int CM_Get_Parent(out uint pdnDevInst, uint dnDevInst, uint ulFlags);

    [DllImport("cfgmgr32.dll", SetLastError = true)]
    private static extern int CM_Get_Device_ID_Size(out uint pulLen, uint dnDevInst, uint ulFlags);

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int CM_Get_Device_IDW(uint dnDevInst, StringBuilder buffer, int bufferLen, uint ulFlags);

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

