namespace DeviceTweakerCS;

internal interface IImodRegisters
{
    bool TryRead(ulong address, out uint value, out string? error);
    bool TryWrite(ulong address, uint value, out string? error);
}

internal static class ImodRegisterService
{
    internal static bool TryWriteInterval(IImodRegisters registers, ulong address, uint interval, out string? error)
    {
        if (!registers.TryRead(address, out uint currentValue, out error)) return false;
        uint mergedValue = (currentValue & 0xFFFF0000) | (interval & 0xFFFF);
        if (!registers.TryWrite(address, mergedValue, out error)) return false;
        if (!registers.TryRead(address, out uint readBack, out string? readError))
        {
            error = $"IMOD write completed, but readback failed: {readError}";
            return false;
        }
        uint expectedInterval = interval & 0xFFFF, actualInterval = readBack & 0xFFFF;
        if (actualInterval != expectedInterval)
        {
            error = $"IMOD readback mismatch at 0x{address:X}: expected interval 0x{expectedInterval:X}, read 0x{actualInterval:X}";
            return false;
        }
        error = null;
        return true;
    }
}
