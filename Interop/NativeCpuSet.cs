using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal static class NativeCpuSet
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct GroupAffinity
    {
        public UIntPtr Mask;
        public ushort Group;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public ushort[] Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SystemCpuSetInformation
    {
        public int Size;
        public int Type;
        public int Id;
        public short Group;
        public byte LogicalProcessorIndex;
        public byte CoreIndex;
        public byte LastLevelCacheIndex;
        public byte NumaNodeIndex;
        public byte EfficiencyClass;
        public byte Parked;
        public byte Allocated;
        public byte AllocatedToTargetProcess;
        public UIntPtr SchedulingClass;
        public UIntPtr AllocationTag;
        public GroupAffinity GroupAffinity;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool GetSystemCpuSetInformation(
        IntPtr information,
        int bufferLength,
        out int returnedLength,
        IntPtr process,
        int flags);
}

