using System.Runtime.InteropServices;

namespace DeviceTweakerCS;

internal static class NativeLogicalProcessor
{
    internal const int RelationProcessorCore = 0;
    internal const int RelationNumaNode = 1;
    internal const int RelationCache = 2;
    internal const int RelationNumaNodeEx = 6;
    internal const int RelationAll = 0xFFFF;

    internal const int RecordHeaderSize = 8;
    internal const int ProcessorGroupCountOffset = 30;
    internal const int ProcessorGroupMasksOffset = 32;
    internal const int NumaNodeNumberOffset = 8;
    internal const int NumaGroupCountOffset = 30;
    internal const int NumaGroupMasksOffset = 32;
    internal const int CacheLevelOffset = 8;
    internal const int CacheTypeOffset = 16;
    internal const int CacheGroupCountOffset = 38;
    internal const int CacheGroupMasksOffset = 40;
    internal static int GroupAffinitySize => IntPtr.Size + 8;

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool GetLogicalProcessorInformationEx(
        int relationshipType,
        IntPtr buffer,
        ref int returnedLength);

    internal static ulong ReadAffinityMask(IntPtr groupAffinity)
    {
        return IntPtr.Size == 8
            ? unchecked((ulong)Marshal.ReadInt64(groupAffinity))
            : unchecked((uint)Marshal.ReadInt32(groupAffinity));
    }
}
