using System.Buffers.Binary;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal sealed record WindowsIrqSnapshot(string Pdo, bool Started, byte[] Raw, byte[] Translated, uint? Support);
internal interface IWindowsIrqSource { WindowsIrqSnapshot Read(string instanceId); }

internal static class WindowsInterruptReader
{
    internal static InterruptModeResult Read(string instanceId)
    {
        if (QaRun.IsReadOnly) return InterruptModeResult.Pending;
        return ReadStable(instanceId, new NativeSource());
    }

    internal static InterruptModeResult ReadStable(string instanceId, IWindowsIrqSource source)
    {
        if (!instanceId.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
            return new(InterruptMode.NotPci, "This device has no PCI interrupt mode.");
        try
        {
            // A changed PDO, resource list or capability mask invalidates the entire sample.
            for (int attempt = 0; attempt < 2; attempt++)
            {
                WindowsIrqSnapshot a = source.Read(instanceId), b = source.Read(instanceId);
                if (!a.Started || !b.Started) return Fail("Device is not started. Windows interrupt allocation is unavailable.");
                if (a.Pdo == b.Pdo && a.Support == b.Support && a.Raw.AsSpan().SequenceEqual(b.Raw)
                    && a.Translated.AsSpan().SequenceEqual(b.Translated))
                    return Decode(b.Raw, b.Translated, b.Support);
            }
            return Fail("Windows interrupt allocation changed while reading. Refresh to retry.");
        }
        catch (Exception ex) when (ex is IOException or Win32Exception or InvalidDataException or UnauthorizedAccessException)
        { return Fail("Windows interrupt allocation: " + ex.Message); }
    }

    private static InterruptModeResult Fail(string detail) => new(InterruptMode.Unavailable, detail, DateTimeOffset.Now);
    private sealed record Irq(bool Message, ushort Count);

    internal static InterruptModeResult Decode(byte[] raw, byte[] translated, uint? support)
    {
        try
        {
            Irq[] a = Parse(raw), b = Parse(translated);
            if (a.Length != b.Length || a.Where((item, i) => item.Message != b[i].Message).Any())
                return Fail("Windows raw and translated interrupt allocations disagree.");
            if (a.Length == 0) return Result(InterruptMode.None, "No interrupt resources allocated.");
            int messages = a.Count(item => item.Message);
            if (messages == 0) return Result(InterruptMode.Line, "Line-based interrupt resources allocated.");
            if (messages != a.Length) return Fail("Windows allocation contains both Line and message interrupts.");
            if (a.Any(item => item.Count == 0)) return Fail("Windows raw interrupt allocation has an invalid message count.");
            if (support is null) return Result(InterruptMode.Message, "Message interrupts allocated; hardware capabilities unavailable.");
            if ((support.Value & ~7u) != 0) return Fail("Unrecognized hardware interrupt capability flags.");
            bool msi = (support & 2) != 0, msix = (support & 4) != 0;
            if (!msi && !msix) return Fail("Windows message allocation conflicts with hardware capabilities.");
            if (a.Length > 1)
            {
                if (!msix || a.Any(item => item.Count != 1)) return Fail("Unexpected Windows MSI-X resource layout.");
                return Result(InterruptMode.MsiX, "Separate message descriptors allocated.");
            }
            if (a[0].Count > 1)
            {
                if (!msi || a[0].Count is not (2 or 4 or 8 or 16 or 32)) return Fail("Unexpected Windows MSI message count.");
                return Result(InterruptMode.Msi, "One descriptor with multiple messages allocated.");
            }
            if (msi && !msix) return Result(InterruptMode.Msi, "Message interrupt allocated; hardware supports MSI, not MSI-X.");
            if (msix && !msi) return Result(InterruptMode.MsiX, "Message interrupt allocated; hardware supports MSI-X, not MSI.");
            return Result(InterruptMode.Message, "One message allocated; hardware supports MSI and MSI-X. These resources cannot distinguish MSI from MSI-X.");
        }
        catch (InvalidDataException ex) { return Fail("Windows interrupt allocation: " + ex.Message); }
    }

    private static InterruptModeResult Result(InterruptMode mode, string reason) => new(mode,
        "Source: Windows allocated resources. " + reason + " PCI Enable bits were not read.", DateTimeOffset.Now);

    private static Irq[] Parse(byte[] data)
    {
        if (data.Length < 4 || data.Length > 1048576) throw new InvalidDataException("Invalid resource list size.");
        uint full = BinaryPrimitives.ReadUInt32LittleEndian(data);
        if (full > 256) throw new InvalidDataException("Invalid full descriptor count.");
        int offset = 4; List<Irq> irqs = [];
        for (uint i = 0; i < full; i++)
        {
            if (data.Length - offset < 16) throw new InvalidDataException("Truncated full descriptor.");
            uint count = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset + 12)); offset += 16;
            if (count > (data.Length - offset) / 20) throw new InvalidDataException("Truncated partial descriptors.");
            uint extra = 0;
            for (uint j = 0; j < count; j++, offset += 20)
            {
                if (data[offset] == 2)
                {
                    ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset + 2));
                    irqs.Add(new((flags & 2) != 0, BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset + 6))));
                }
                if (data[offset] == 5)
                {
                    if (j != count - 1) throw new InvalidDataException("Misplaced device-specific data.");
                    extra = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset + 4));
                }
            }
            if (extra > data.Length - offset) throw new InvalidDataException("Truncated device-specific data.");
            offset += (int)extra;
        }
        if (offset != data.Length) throw new InvalidDataException("Unexpected resource list trailing bytes.");
        return irqs.ToArray();
    }

    private sealed class NativeSource : IWindowsIrqSource
    {
        public WindowsIrqSnapshot Read(string instanceId)
        {
            if (QaRun.IsReadOnly) throw new InvalidOperationException("QA native Windows resource read blocked.");
            if (!Environment.Is64BitProcess || RuntimeInformation.ProcessArchitecture != Architecture.X64)
                throw new InvalidDataException("Windows resource reader requires native x64.");
            if (!NativeCfgMgr32.TryGetPciPdoName(instanceId, out string pdo, out string error)) throw new IOException(error);
            if (!pdo.StartsWith(@"\Device\", StringComparison.OrdinalIgnoreCase) || pdo.Length > 255 || pdo.Contains('\0'))
                throw new InvalidDataException("Invalid device PDO name.");
            int cr = CM_Locate_DevNodeW(out uint node, instanceId, 0);
            if (cr != 0) throw new IOException($"CM_Locate_DevNode: 0x{cr:X}");
            cr = CM_Get_DevNode_Status(out uint status, out uint problem, node, 0);
            if (cr != 0) throw new IOException($"CM_Get_DevNode_Status: 0x{cr:X}");
            bool started = (status & 8) != 0 && problem == 0;
            if (!started) return new(pdo, false, [], [], null);
            uint? support = NativeCfgMgr32.TryGetPciInterruptCapabilities(instanceId, out var capabilities)
                ? (uint?)capabilities!.Support : null;
            using var key = OpenMap();
            return new(pdo, true, ReadValue(key, pdo + ".Raw"), ReadValue(key, pdo + ".Translated"), support);
        }

        private static Microsoft.Win32.SafeHandles.SafeRegistryHandle OpenMap()
        {
            int error = RegOpenKeyEx(new IntPtr(unchecked((int)0x80000002)), @"HARDWARE\RESOURCEMAP\PnP Manager\PnpManager",
                0, 0x20119, out var key); // KEY_READ | KEY_WOW64_64KEY
            if (error != 0) { key.Dispose(); throw new Win32Exception(error); }
            return key;
        }

        private static byte[] ReadValue(Microsoft.Win32.SafeHandles.SafeRegistryHandle key, string name)
        {
            uint size = 0;
            int error = RegQueryValueEx(key, name, IntPtr.Zero, out uint type, null, ref size);
            if (error != 0) throw new Win32Exception(error, "Allocated resource map unavailable for this device.");
            if (type != 8 || size < 4 || size > 1048576) throw new InvalidDataException("Unexpected registry resource type or size.");
            byte[] data = new byte[size];
            error = RegQueryValueEx(key, name, IntPtr.Zero, out type, data, ref size);
            if (error != 0) throw new Win32Exception(error);
            if (type != 8 || size != data.Length) throw new InvalidDataException("Resource value changed while reading.");
            return data;
        }
    }

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode)] private static extern int CM_Locate_DevNodeW(out uint node, string id, uint flags);
    [DllImport("cfgmgr32.dll")] private static extern int CM_Get_DevNode_Status(out uint status, out uint problem, uint node, uint flags);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] private static extern int RegOpenKeyEx(IntPtr root, string subkey, uint options, uint access, out Microsoft.Win32.SafeHandles.SafeRegistryHandle key);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] private static extern int RegQueryValueEx(Microsoft.Win32.SafeHandles.SafeRegistryHandle key, string name, IntPtr reserved, out uint type, byte[]? bytes, ref uint size);
}
