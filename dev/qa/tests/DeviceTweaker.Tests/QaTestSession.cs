using System.Runtime.CompilerServices;
using DeviceTweakerCS;
using Xunit;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace DeviceTweaker.Tests;

internal static class QaTestSession
{
    [ModuleInitializer]
    internal static void Initialize()
    {
        Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY", "1");
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_OUTPUT")))
            Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_OUTPUT",
                Path.Combine(Path.GetTempPath(), "DEVICE_TWEAKER_QA", $"tests-{Environment.ProcessId}"));
        _ = QaRun.IsReadOnly;
    }
}
