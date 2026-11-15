namespace DeviceTweakerCS;

internal static class ImodReadbackService
{
    // The displayed index is the array index. A failed register must not shift
    // later values into the wrong interrupter rows or a role map.
    internal static bool TryRead(IDeviceRegisters registers, ulong runtimeAddress, uint count,
        out List<uint> values, out string? error)
    {
        values = [];
        error = null;
        ulong lastOffset = 0x24UL + 0x20UL * (count == 0 ? 0 : count - 1);
        if (count == 0 || count > 2048 || runtimeAddress > ulong.MaxValue - lastOffset - 3)
        {
            error = "invalid IMOD readback range";
            return false;
        }

        List<uint> complete = new((int)count);
        for (uint i = 0; i < count; i++)
        {
            ulong address = runtimeAddress + 0x24UL + 0x20UL * i;
            if (!DeviceTweakerImodRuntime.TryRead32(registers, address, out uint raw, out string readError))
            {
                error = $"interrupter={i} address=0x{address:X}: {readError}";
                return false;
            }
            complete.Add(raw & 0xFFFF);
        }
        values = complete;
        return true;
    }
}
