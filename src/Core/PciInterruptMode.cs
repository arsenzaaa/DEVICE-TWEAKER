using System.Buffers.Binary;

namespace DeviceTweakerCS;

internal enum InterruptMode { NotRead, Line, Msi, MsiX, Message, None, Unavailable, NotPci }

internal sealed record InterruptModeResult(InterruptMode Mode, string Detail, DateTimeOffset? ReadAt = null)
{
    public string Text => Mode switch
    {
        InterruptMode.Line => "Line", InterruptMode.Msi => "MSI", InterruptMode.MsiX => "MSI-X",
        InterruptMode.Message => "MSI/MSI-X", InterruptMode.None => "no IRQ",
        InterruptMode.NotRead => "not read", InterruptMode.NotPci => "not PCI", _ => "read failed",
    };
    public static InterruptModeResult Pending { get; } = new(InterruptMode.NotRead, "Reading interrupt mode.");
}

/// <summary>Uses allocated resources, not boot requirements, WMI numbers or MSISupported.</summary>
internal static class PciInterruptMode
{
    internal const int ReplyHeaderSize = 276, ResourceMaximum = 65536, ReplySize = ReplyHeaderSize + ResourceMaximum;
    private static uint U32(ReadOnlySpan<byte> data, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(data[offset..]);
    private static ushort U16(ReadOnlySpan<byte> data, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(data[offset..]);

    internal static InterruptModeResult DecodeReply(string instanceId, ReadOnlySpan<byte> reply)
    {
        InterruptModeResult Fail(string detail) => new(InterruptMode.Unavailable, detail, DateTimeOffset.Now);
        if (reply.Length < ReplyHeaderSize || U32(reply, 0) != 1) return Fail("DTIMOD PCI query: unsupported / incomplete reply");
        uint configStatus = U32(reply, 4), resourcesStatus = U32(reply, 8);
        uint configBytes = U32(reply, 12), resourceBytes = U32(reply, 16);
        if (configBytes > 256 || resourceBytes > ResourceMaximum || resourceBytes > reply.Length - ReplyHeaderSize)
            return Fail("DTIMOD PCI query: invalid reply lengths");
        if (resourcesStatus != 0) return Fail($"Windows allocated resources: NTSTATUS 0x{resourcesStatus:X8}");
        if (!TryReadAllocatedInterrupts(reply.Slice(ReplyHeaderSize, (int)resourceBytes), out bool line, out bool message, out string error))
            return Fail(error);
        if (line && message) return Fail("Windows allocated resources contain both Line and message interrupts");
        if (!line && !message) return new(InterruptMode.None, "Windows has allocated no interrupt resources to this device.", DateTimeOffset.Now);
        if (configStatus != 0 || configBytes != 256)
        {
            string detail = $"Windows allocation: {(line ? "Line" : "message interrupts")}. PCI read: NTSTATUS 0x{configStatus:X8}, bytes={configBytes}.";
            return new(line ? InterruptMode.Line : InterruptMode.Message, detail, DateTimeOffset.Now);
        }
        if (!TryReadPci(reply.Slice(20, 256), instanceId, out bool msi, out bool msiX, out error)) return Fail(error);
        if ((msi && msiX) || (line && (msi || msiX)) || (message && !msi && !msiX))
            return Fail($"PCI / Windows interrupt state disagrees: Line={line}, Message={message}, MSI Enable={msi}, MSI-X Enable={msiX}. Refresh IRQ Mode to retry.");
        InterruptMode mode = line ? InterruptMode.Line : msiX ? InterruptMode.MsiX : InterruptMode.Msi;
        return new(mode, $"Windows allocated resources / PCI Enable bits: {mode}. Exact device instance verified by DTIMOD.", DateTimeOffset.Now);
    }

    internal static bool TryReadPci(ReadOnlySpan<byte> config, string instanceId, out bool msi, out bool msiX, out string error)
    {
        msi = msiX = false; error = string.Empty;
        if (config.Length != 256) { error = "PCI configuration must contain 256 bytes"; return false; }
        ushort vendor = U16(config, 0), device = U16(config, 2);
        if (vendor is 0 or 0xFFFF || !instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase)
            || !instanceId.Contains($"VEN_{vendor:X4}&DEV_{device:X4}", StringComparison.OrdinalIgnoreCase))
        { error = "PCI vendor / device does not match the requested instance"; return false; }
        int headerType = config[0x0E] & 0x7F;
        if (headerType is not (0 or 1)) { error = "PCI header type is not supported"; return false; }
        if ((U16(config, 6) & 0x10) == 0) return true;
        int offset = config[0x34];
        HashSet<int> visited = [];
        bool foundMsi = false, foundMsiX = false;
        while (offset != 0)
        {
            if (offset < 0x40 || offset > 0xFC || (offset & 3) != 0 || !visited.Add(offset))
            { error = "PCI capability list is invalid / cyclic"; return false; }
            byte id = config[offset];
            if (id is 0 or 0xFF) { error = "PCI capability ID is invalid"; return false; }
            if (id == 0x05)
            {
                if (foundMsi) { error = "Duplicate PCI MSI capability"; return false; }
                foundMsi = true;
                msi = (U16(config, offset + 2) & 1) != 0;
            }
            if (id == 0x11)
            {
                if (foundMsiX) { error = "Duplicate PCI MSI-X capability"; return false; }
                foundMsiX = true;
                msiX = (U16(config, offset + 2) & 0x8000) != 0;
            }
            offset = config[offset + 1];
        }
        return true;
    }

    // Native x64 WDK layout is compile-time asserted by imod_driver.c.
    internal static bool TryReadAllocatedInterrupts(ReadOnlySpan<byte> data, out bool line, out bool message, out string error)
    {
        line = message = false;
        if (!TryReadResourceList(data, out int consumed, out bool rawLine, out bool rawMessage, out error)) return false;
        if (consumed != data.Length)
        {
            // Windows 10 DevicePropertyAllocatedResources returns raw and translated lists consecutively.
            // Parse the entire second list, rather than discarding bytes after the first one.
            if (!TryReadResourceList(data[consumed..], out int translatedBytes, out bool translatedLine, out bool translatedMessage, out error)) return false;
            if (translatedBytes != data.Length - consumed)
            { error = "Windows allocated resources: unexpected trailing bytes"; return false; }
            if (rawLine != translatedLine || rawMessage != translatedMessage)
            { error = "Windows allocated resources: raw / translated interrupt types disagree"; return false; }
        }
        line = rawLine; message = rawMessage;
        return true;
    }

    private static bool TryReadResourceList(ReadOnlySpan<byte> data, out int consumed, out bool line, out bool message, out string error)
    {
        consumed = 0;
        line = message = false; error = string.Empty;
        if (data.Length < 4) { error = "Windows allocated resources: incomplete CM_RESOURCE_LIST"; return false; }
        uint fullCount = U32(data, 0);
        int offset = 4;
        for (uint full = 0; full < fullCount; full++)
        {
            if (data.Length - offset < 16) { error = "Windows allocated resources: incomplete full descriptor"; return false; }
            uint count = U32(data, offset + 12);
            offset += 16;
            if (count > (data.Length - offset) / 20) { error = "Windows allocated resources: invalid descriptor count"; return false; }
            uint extraBytes = 0;
            for (uint part = 0; part < count; part++, offset += 20)
            {
                byte type = data[offset];
                if (type == 2) // CmResourceTypeInterrupt / CM_RESOURCE_INTERRUPT_MESSAGE = 2
                { if ((U16(data, offset + 2) & 2) != 0) message = true; else line = true; }
                if (type == 5) // DeviceSpecific must be last; bytes follow all fixed descriptors.
                {
                    if (part != count - 1) { error = "Windows allocated resources: misplaced device-specific data"; return false; }
                    extraBytes = U32(data, offset + 4);
                }
            }
            if (extraBytes > data.Length - offset) { error = "Windows allocated resources: truncated device-specific data"; return false; }
            offset += (int)extraBytes;
        }
        consumed = offset;
        return true;
    }
}
