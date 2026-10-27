#if !DEVELOPER_QA
namespace DeviceTweakerCS;
public sealed partial class MainForm
{
    private static TestDeviceState EnsureTestDeviceState(DeviceInfo device)
        => throw new InvalidOperationException("Device simulation requires the developer QA build.");
    private static bool ApplyTestBlockSettings(DeviceBlock block, bool autoMsiOnly, OperationReport? report)
        => throw new InvalidOperationException("Device simulation requires the developer QA build.");
    private static void ResetTestBlockState(DeviceBlock block)
        => throw new InvalidOperationException("Device simulation requires the developer QA build.");
    private static void RefreshTestImodPreview(DeviceBlock block, string reason = "test-preview")
        => throw new InvalidOperationException("Device simulation requires the developer QA build.");
}
#endif
