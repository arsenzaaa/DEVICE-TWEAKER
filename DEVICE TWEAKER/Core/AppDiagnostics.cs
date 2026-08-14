using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace DeviceTweakerCS;

internal static class AppDiagnostics
{
    private const string LogFolderName = "logs";
    private static readonly object Sync = new();
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);
    private static readonly UTF8Encoding Utf8WithBom = new(encoderShouldEmitUTF8Identifier: true);
    private static readonly long SessionStartTimestamp = Stopwatch.GetTimestamp();

    private static bool _enabled;
    private static string? _sessionLogPath;
    private static long _sequence;

    internal static string LogDirectory => Path.Combine(
        AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar),
        LogFolderName);

    internal static string? SessionLogPath
    {
        get
        {
            lock (Sync)
            {
                return _sessionLogPath;
            }
        }
    }

    internal static bool TryEnable(out string? path, out string? error)
    {
        lock (Sync)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_sessionLogPath))
                {
                    Directory.CreateDirectory(LogDirectory);
                    string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                    string fileName = $"DeviceTweaker_{stamp}.log";
                    string candidate = Path.Combine(LogDirectory, fileName);

                    using FileStream stream = new(
                        candidate,
                        FileMode.CreateNew,
                        FileAccess.Write,
                        FileShare.ReadWrite);
                    byte[] preamble = Utf8WithBom.GetPreamble();
                    stream.Write(preamble, 0, preamble.Length);

                    _sessionLogPath = candidate;
                    _sequence = 0;
                }

                _enabled = true;
                path = _sessionLogPath;
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                _enabled = false;
                path = null;
                error = ex.ToString();
                Debug.WriteLine($"DEVICE TWEAKER logging initialization failed: {ex}");
                return false;
            }
        }
    }

    internal static void Disable()
    {
        lock (Sync)
        {
            _enabled = false;
        }
    }

    internal static bool Write(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        lock (Sync)
        {
            if (!_enabled || string.IsNullOrWhiteSpace(_sessionLogPath))
            {
                return false;
            }

            try
            {
                AppendSessionLineUnsafe(message);
                return true;
            }
            catch (Exception ex)
            {
                _enabled = false;
                Debug.WriteLine($"DEVICE TWEAKER log write failed: {ex}");
                return false;
            }
        }
    }

    internal static void WriteFatal(string source, Exception ex)
    {
        lock (Sync)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                string crashName = $"crash_{stamp}_PID{Environment.ProcessId}.txt";
                string crashPath = Path.Combine(LogDirectory, crashName);
                string latestPath = Path.Combine(LogDirectory, "last-crash.txt");

                StringBuilder report = new();
                report.AppendLine(DateTime.Now.ToString("O"));
                report.AppendLine($"source={source}");
                report.AppendLine($"version={GetAppVersion()}");
                report.AppendLine($"pid={Environment.ProcessId}");
                report.AppendLine($"processArchitecture={RuntimeInformation.ProcessArchitecture}");
                report.AppendLine($"osArchitecture={RuntimeInformation.OSArchitecture}");
                report.AppendLine($"dotnet={Environment.Version}");
                report.AppendLine($"sessionLog={_sessionLogPath ?? "unavailable"}");
                report.AppendLine();
                report.AppendLine(ex.ToString());

                string text = report.ToString();
                File.WriteAllText(crashPath, text, Utf8WithBom);
                File.WriteAllText(latestPath, text, Utf8WithBom);

                if (!string.IsNullOrWhiteSpace(_sessionLogPath))
                {
                    string message = Flatten(
                        $"FATAL: source={source} type={ex.GetType().FullName} message=\"{ex.Message}\" crashFile=\"{crashPath}\"");
                    AppendSessionLineUnsafe(message);
                }
            }
            catch (Exception loggingError)
            {
                Debug.WriteLine($"DEVICE TWEAKER crash logging failed: {loggingError}");
            }
        }
    }

    private static void AppendSessionLineUnsafe(string message)
    {
        if (string.IsNullOrWhiteSpace(_sessionLogPath))
        {
            return;
        }

        long sequence = ++_sequence;
        TimeSpan elapsed = Stopwatch.GetElapsedTime(SessionStartTimestamp);
        string line =
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] " +
            $"[#{sequence:D6}] [+{elapsed.TotalMilliseconds:0}ms] [T{Environment.CurrentManagedThreadId:D2}] " +
            message;
        File.AppendAllText(_sessionLogPath, line + Environment.NewLine, Utf8NoBom);
    }

    private static string Flatten(string value)
    {
        return value
            .Replace("\r\n", " | ")
            .Replace("\n", " | ")
            .Replace("\r", " | ")
            .Trim();
    }

    private static string GetAppVersion()
    {
        try
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            return assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                ?? assembly.GetName().Version?.ToString()
                ?? "unknown";
        }
        catch
        {
            return "unknown";
        }
    }
}
