using System.Diagnostics;
using static DeviceTweakerCS.ImodDriverNative;
using static DeviceTweakerCS.ImodDriverFiles;

namespace DeviceTweakerCS;

// KDU process lifetime, captured output and the existing uncertain-outcome gate.
// File staging belongs to ImodDriverFiles; this class never owns a form/control.
internal static class KduDriverLoader
{
    internal static bool TryLoadImodDriverWithKduFallback(string driverPath, Action<string>? log, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryLoadImodDriverWithKduFallback");
        error = null;

        if (IsImodKduFallbackDisabled())
        {
            error = $"{ImodKduDisableEnv}=1";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason={error}");
            return false;
        }

        if (!File.Exists(driverPath))
        {
            error = $"driver missing: {driverPath}";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason=driver_missing path={driverPath}");
            return false;
        }

        if (!EnsureImodKduPayloadOnDisk(log, out string kduPath, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: payload unavailable error={CompactLogValue(error)}");
            return false;
        }

        string kduDirectory = Path.GetDirectoryName(kduPath) ?? string.Empty;
        string kduDatabasePath = Path.Combine(kduDirectory, ImodKduDatabaseFileName);
        if (!File.Exists(kduDatabasePath))
        {
            error = $"drv64.dll missing next to kdu.exe: {kduDatabasePath}";
            log?.Invoke($"IMOD.DRIVER.KDU: skipped reason=db_missing path={kduDatabasePath}");
            return false;
        }

        if (!TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduPath, out error)
            || !TryVerifyEmbeddedImodResourceFile("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", kduDatabasePath, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: payload verification failed error={CompactLogValue(error)}");
            return false;
        }

        log?.Invoke($"IMOD.DRIVER.KDU: map start kdu={kduPath} driver={driverPath}");
        if (!RunImodKduMap(kduPath, driverPath, log, out error))
        {
            log?.Invoke($"IMOD.DRIVER.KDU: map failed error={CompactLogValue(error)}");
            return false;
        }

        if (TryOpenImodDriverDevice(out IntPtr handle, out string? openError))
        {
            _ = CloseHandle(handle);
            log?.Invoke("IMOD.DRIVER.KDU: device available");
            return true;
        }

        error = $"mapped but device unavailable: {openError}";
        log?.Invoke($"IMOD.DRIVER.KDU: map result unusable error={CompactLogValue(error)}");
        return false;
    }

    private static bool IsImodKduFallbackDisabled()
    {
        string? value = Environment.GetEnvironmentVariable(ImodKduDisableEnv);
        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool RunImodKduMap(string kduPath, string driverPath, Action<string>? log, out string? error)
    {
        QaRun.DemandWrite("RunImodKduMap");
        error = null;
        Stopwatch timer = Stopwatch.StartNew();

        try
        {
            using Process process = new();
            process.StartInfo = BuildStartInfo(kduPath, driverPath);

            DriverLoadGate.BeginLoader();
            bool started;
            try { started = process.Start(); }
            catch { DriverLoadGate.CompleteLoader(); throw; }
            if (!started)
            {
                DriverLoadGate.CompleteLoader();
                error = "failed to start kdu.exe";
                return false;
            }

            Task<string> stdout = process.StandardOutput.ReadToEndAsync();
            Task<string> stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(ImodKduTimeoutMs))
            {
                _ = Task.WaitAll(new Task[] { stdout, stderr }, 2000);
                string timeoutOutput = stdout.IsCompletedSuccessfully ? stdout.Result : string.Empty;
                string timeoutError = stderr.IsCompletedSuccessfully ? stderr.Result : string.Empty;
                log?.Invoke($"IMOD.DRIVER.KDU: timeout elapsedMs={timer.ElapsedMilliseconds}");
                LogImodKduOutput(log, "STDOUT", timeoutOutput);
                LogImodKduOutput(log, "STDERR", timeoutError);
                error = "KDU has not finished. Restart Windows before loading DTIMOD again. See the session log.";
                return false;
            }

            DriverLoadGate.CompleteLoader();

            _ = Task.WaitAll(new Task[] { stdout, stderr }, 2000);
            string output = stdout.IsCompletedSuccessfully ? stdout.Result : string.Empty;
            string errOutput = stderr.IsCompletedSuccessfully ? stderr.Result : string.Empty;
            log?.Invoke($"IMOD.DRIVER.KDU: exit code={process.ExitCode} elapsedMs={timer.ElapsedMilliseconds}");
            LogImodKduOutput(log, "STDOUT", output);
            LogImodKduOutput(log, "STDERR", errOutput);

            if (TryOpenImodDriverDevice(out IntPtr mappedHandle, out string? mappedOpenError))
            {
                _ = CloseHandle(mappedHandle);
                if (process.ExitCode != 0)
                {
                    log?.Invoke(
                        "IMOD.DRIVER.KDU: non-zero exit ignored because device is available "
                        + $"code={process.ExitCode}");
                }

                return true;
            }

            if (process.ExitCode != 0)
            {
                error = $"kdu.exe failed code={process.ExitCode}; DTIMOD device unavailable: {mappedOpenError}. See full KDU output in the session log.";
                return false;
            }

            error = $"kdu.exe exited successfully but DTIMOD device was not created: {mappedOpenError}. See full KDU output in the session log.";
            return false;
        }
        catch (Exception ex)
        {
            log?.Invoke($"IMOD.DRIVER.KDU.EXCEPTION: {LogText.Flatten(ex.ToString())}");
            error = $"{ex.GetType().Name}: {ex.Message}";
            return false;
        }
    }

    private static void LogImodKduOutput(Action<string>? log, string streamName, string? content)
    {
        if (log is null)
        {
            return;
        }

        string normalized = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        log($"IMOD.DRIVER.KDU.{streamName}: BEGIN chars={normalized.Length} lines={lines.Length}");
        if (normalized.Length == 0)
        {
            log($"IMOD.DRIVER.KDU.{streamName}: <empty>");
        }
        else
        {
            for (int i = 0; i < lines.Length; i++)
            {
                log($"IMOD.DRIVER.KDU.{streamName}: {i + 1:D4}: {lines[i]}");
            }
        }
        log($"IMOD.DRIVER.KDU.{streamName}: END");
    }

    private static string QuoteProcessArgument(string value)
    {
        return "\"" + value.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
    }

    internal static ProcessStartInfo BuildStartInfo(string kduPath, string driverPath)
    {
        return new ProcessStartInfo
        {
            FileName = kduPath,
            Arguments = $"-scv 3 -drvn DeviceTweakerImod2 -drvr DeviceTweakerImod2 -map {QuoteProcessArgument(driverPath)}",
            WorkingDirectory = Path.GetDirectoryName(kduPath) ?? AppContext.BaseDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
    }

    private static string CompactLogValue(string? value)
        => string.IsNullOrWhiteSpace(value) ? "none" : value.Replace('\r', ' ').Replace('\n', ' ').Trim();
}
