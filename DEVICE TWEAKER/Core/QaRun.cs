using System.Collections.Concurrent;
using System.Text.Json;
using System.Reflection;
using System.Security.Cryptography;

namespace DeviceTweakerCS;

// Frozen at process startup. TEST ADMIN can change the virtual view, never this policy.
internal static class QaRun
{
    static QaRun() { }
    internal static readonly bool IsReadOnly =
#if DEVELOPER_QA
        true;
#else
        Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY") == "1";
#endif
    internal static readonly string SessionId =
        Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_SESSION") ?? Guid.NewGuid().ToString("N");
    internal static readonly string OutputDirectory = Path.GetFullPath(
        Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_OUTPUT")
        ?? Path.Combine(Path.GetTempPath(), "DEVICE_TWEAKER_QA", SessionId));
    private static readonly ConcurrentQueue<BlockedOperation> Blocked = new();
    private static readonly object ReportSync = new();
    private static readonly Lazy<object> BuildIdentity = new(() =>
    {
        Assembly assembly = typeof(QaRun).Assembly;
        using Stream payload = assembly.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys")
            ?? throw new InvalidDataException("Embedded DTIMOD payload is missing.");
        return new
        {
            version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            assemblyName = assembly.GetName().Name,
            appBuildId = AppBuildIdentity.BuildId,
            executableSha256 = AppBuildIdentity.ExecutableSha256,
            hasTestPanel = typeof(MainForm).GetMethod("ShowTestAdminDialog", BindingFlags.NonPublic | BindingFlags.Instance) is not null,
            embeddedResources = assembly.GetManifestResourceNames(),
            driverSha256 = Convert.ToHexString(SHA256.HashData(payload))
        };
    });

    internal sealed record BlockedOperation(string Operation, DateTimeOffset Time);
    internal static int BlockedCount => Blocked.Count;
    internal static BlockedOperation[] BlockedOperations => Blocked.ToArray();

    internal static bool TryBlock(string operation)
    {
        if (!IsReadOnly) return false;
        Blocked.Enqueue(new(operation, DateTimeOffset.UtcNow));
        AppDiagnostics.Write($"QA.WRITE.BLOCKED: operation={operation}");
        WriteReport();
        return true;
    }

    internal static void DemandWrite(string operation)
    {
        if (TryBlock(operation)) throw new QaWriteBlockedException(operation);
    }

    internal static void DemandArtifactPath(string path)
    {
        if (!IsReadOnly) return;
        string fullPath = Path.GetFullPath(path);
        string prefix = OutputDirectory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            DemandWrite("FILE OUTSIDE QA OUTPUT");
    }

    internal static void WriteReport()
    {
        if (!IsReadOnly) return;
        lock (ReportSync)
        {
            Directory.CreateDirectory(OutputDirectory);
            string target = Path.Combine(OutputDirectory, $"process-{Environment.ProcessId}.json");
            string temporary = target + ".tmp";
            File.WriteAllText(temporary,
                JsonSerializer.Serialize(new
                {
                    schemaVersion = 1,
                    processId = Environment.ProcessId,
                    sessionId = SessionId,
                    readOnly = IsReadOnly,
                    desktop = new { expected = QaDesktop.Expected, actual = QaDesktop.CurrentName(),
                        isolated = QaDesktop.Expected is not null && QaDesktop.CurrentName() == QaDesktop.Expected },
                    build = BuildIdentity.Value,
                    logPath = AppDiagnostics.SessionLogPath,
                    blockedOperations = Blocked.ToArray(),
                    validationScope = "Managed write guards and virtual scenarios; hardware writes/readback are not executed.",
                    hardwareValidation = "SKIPPED"
                }, new JsonSerializerOptions { WriteIndented = true }));
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    if (File.Exists(target)) File.Replace(temporary, target, destinationBackupFileName: null);
                    else File.Move(temporary, target);
                    break;
                }
                catch (Exception ex) when (attempt < 10 && ex is IOException or UnauthorizedAccessException)
                {
                    // Windows can briefly retain a previous reader's delete-sharing
                    // handle during replacement. Never fall back to truncating target.
                    Thread.Sleep(5);
                }
            }
        }
    }
}

internal sealed class QaWriteBlockedException(string operation)
    : InvalidOperationException($"QA read-only session blocked: {operation}.");
