using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal static class NativeCpuSet
{
    // SYSTEM_CPU_SET_INFORMATION_TYPE.CpuSetInformation.
    // Keep this layout byte-for-byte compatible with winnt.h. The CpuSet
    // member of the native union is 24 bytes and the complete record is
    // 32 bytes on both x86 and x64.
    [StructLayout(LayoutKind.Sequential)]
    internal struct SystemCpuSetInformation
    {
        public uint Size;
        public uint Type;
        public uint Id;
        public ushort Group;
        public byte LogicalProcessorIndex;
        public byte CoreIndex;
        public byte LastLevelCacheIndex;
        public byte NumaNodeIndex;
        public byte EfficiencyClass;
        public byte AllFlags;
        public byte SchedulingClass;
        private readonly byte Reserved;
        public ulong AllocationTag;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool GetSystemCpuSetInformation(
        IntPtr information,
        int bufferLength,
        out int returnedLength,
        IntPtr process,
        int flags);
}
