using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal static partial class AppDiagnostics
{
    private const string LogFolderName = "logs";
    private const int RetainedSessionLogCount = 30;
    private const int RetainedCrashReportCount = 10;
    private const int LogSchemaVersion = 3;
    private static readonly object Sync = new();
    private static readonly UTF8Encoding Utf8WithBom = new(encoderShouldEmitUTF8Identifier: true);
    private static readonly long SessionStartTimestamp = Stopwatch.GetTimestamp();
    private static readonly Dictionary<string, long> CategoryCounts = new(StringComparer.OrdinalIgnoreCase);

    private static bool _enabled;
    private static string? _sessionLogPath;
    private static StreamWriter? _sessionWriter;
    private static long _sequence;
    private static long _infoCount;
    private static long _warningCount;
    private static long _errorCount;
    private static long _fatalCount;
    private static long _continuationCount;

    private enum DiagnosticLevel
    {
        Info,
        Warning,
        Error,
        Fatal,
    }

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

    internal static int SchemaVersion => LogSchemaVersion;

    internal static bool TryEnable(out string? path, out string? error)
    {
        lock (Sync)
        {
            try
            {
                if (_sessionWriter is null)
                {
                    Directory.CreateDirectory(LogDirectory);
                    PruneLogsUnsafe();
                    string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                    string candidate = Path.Combine(LogDirectory, $"DeviceTweaker_{stamp}.log");

                    FileStream stream = new(
                        candidate,
                        FileMode.CreateNew,
                        FileAccess.Write,
                        FileShare.ReadWrite,
                        bufferSize: 16 * 1024,
                        FileOptions.SequentialScan);
                    _sessionWriter = new StreamWriter(stream, Utf8WithBom, bufferSize: 16 * 1024)
                    {
                        AutoFlush = true,
                    };

                    _sessionLogPath = candidate;
                    ResetSessionCountersUnsafe();
                }

                _enabled = true;
                path = _sessionLogPath;
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                CloseWriterUnsafe();
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
            CloseWriterUnsafe();
        }
    }

    internal static void CompleteSession(string reason)
    {
        lock (Sync)
        {
            if (!_enabled || _sessionWriter is null)
            {
                CloseWriterUnsafe();
                _enabled = false;
                return;
            }

            try
            {
                long entriesBeforeSummary = _sequence;
                TimeSpan elapsed = Stopwatch.GetElapsedTime(SessionStartTimestamp);
                string categories = string.Join(
                    ',',
                    CategoryCounts
                        .OrderByDescending(pair => pair.Value)
                        .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(pair => $"{pair.Key}:{pair.Value}"));
                AppendMessageUnsafe(
                    $"LOG.SESSION.SUMMARY: status=completed reason=\"{Flatten(reason)}\" durationMs={elapsed.TotalMilliseconds:0} " +
                    $"entries={entriesBeforeSummary} infoEntries={_infoCount} warningEntries={_warningCount} errorEntries={_errorCount} " +
                    $"fatalEntries={_fatalCount} continuationEntries={_continuationCount} categories=[{categories}]",
                    DiagnosticLevel.Info,
                    "LOG");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEVICE TWEAKER session summary failed: {ex}");
            }
            finally
            {
                _enabled = false;
                CloseWriterUnsafe();
            }
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
            if (!_enabled || _sessionWriter is null)
            {
                return false;
            }

            try
            {
                string normalized = message.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
                string[] lines = normalized.Split('\n');
                int firstIndex = Array.FindIndex(lines, line => !string.IsNullOrWhiteSpace(line));
                if (firstIndex < 0)
                {
                    return false;
                }

                string firstLine = lines[firstIndex].Trim();
                string category = InferCategory(firstLine);
                DiagnosticLevel level = InferLevel(firstLine);
                AppendMessageUnsafe(firstLine, level, category);

                for (int index = firstIndex + 1; index < lines.Length; index++)
                {
                    string continuation = lines[index].Trim();
                    if (continuation.Length == 0)
                    {
                        continue;
                    }

                    _continuationCount++;
                    AppendMessageUnsafe($"CONT | {continuation}", level, category);
                }

                return true;
            }
            catch (Exception ex)
            {
                _enabled = false;
                CloseWriterUnsafe();
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
                PruneCrashReportsUnsafe();
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                string crashPath = Path.Combine(LogDirectory, $"crash_{stamp}_PID{Environment.ProcessId}.txt");
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

                if (_enabled && _sessionWriter is not null)
                {
                    AppendMessageUnsafe(
                        $"FATAL: source={source} type={ex.GetType().FullName} message=\"{Flatten(ex.Message)}\" crashFile=\"{crashPath}\"",
                        DiagnosticLevel.Fatal,
                        "FATAL");
                }
            }
            catch (Exception loggingError)
            {
                Debug.WriteLine($"DEVICE TWEAKER crash logging failed: {loggingError}");
            }
        }
    }

    private static void ResetSessionCountersUnsafe()
    {
        _sequence = 0;
        _infoCount = 0;
        _warningCount = 0;
        _errorCount = 0;
        _fatalCount = 0;
        _continuationCount = 0;
        CategoryCounts.Clear();
    }

    private static void PruneLogsUnsafe()
    {
        try
        {
            PrunePatternUnsafe("DeviceTweaker_*.log", RetainedSessionLogCount);
            PrunePatternUnsafe("imod_startup_*.log", RetainedSessionLogCount);
            PruneCrashReportsUnsafe();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"DEVICE TWEAKER log pruning failed: {ex}");
        }
    }

    private static void PruneCrashReportsUnsafe()
    {
        PrunePatternUnsafe("crash_*.txt", RetainedCrashReportCount);
    }

    private static void PrunePatternUnsafe(string pattern, int retainedCount)
    {
        if (!Directory.Exists(LogDirectory))
        {
            return;
        }

        FileInfo[] files = new DirectoryInfo(LogDirectory)
            .EnumerateFiles(pattern, SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .ToArray();

        foreach (FileInfo file in files.Skip(Math.Max(1, retainedCount)))
        {
            file.Delete();
        }
    }

    private static void AppendMessageUnsafe(string message, DiagnosticLevel level, string category)
    {
        if (_sessionWriter is null)
        {
            return;
        }

        long sequence = ++_sequence;
        TimeSpan elapsed = Stopwatch.GetElapsedTime(SessionStartTimestamp);
        string levelText = level switch
        {
            DiagnosticLevel.Warning => "WARN",
            DiagnosticLevel.Error => "ERROR",
            DiagnosticLevel.Fatal => "FATAL",
            _ => "INFO",
        };
        string safeCategory = string.IsNullOrWhiteSpace(category) ? "GENERAL" : category.ToUpperInvariant();
        if (safeCategory.Length > 10)
        {
            safeCategory = safeCategory[..10];
        }

        string line =
            $"[{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz}] " +
            $"[#{sequence:D6}] [+{elapsed.TotalMilliseconds:0000000}ms] [T{Environment.CurrentManagedThreadId:D2}] " +
            $"[{levelText}] [{safeCategory}] {message}";
        _sessionWriter.WriteLine(line);

        switch (level)
        {
            case DiagnosticLevel.Warning:
                _warningCount++;
                break;
            case DiagnosticLevel.Error:
                _errorCount++;
                break;
            case DiagnosticLevel.Fatal:
                _fatalCount++;
                break;
            default:
                _infoCount++;
                break;
        }

        CategoryCounts[safeCategory] = CategoryCounts.TryGetValue(safeCategory, out long count) ? count + 1 : 1;
    }

    private static DiagnosticLevel InferLevel(string message)
    {
        string eventName = GetEventName(message);
        string[] eventParts = eventName.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (eventName.StartsWith("TEST.", StringComparison.OrdinalIgnoreCase)
            && TestPassMessageRegex().IsMatch(message))
        {
            return DiagnosticLevel.Info;
        }

        // A missing DTIMOD device before its loader runs is expected state,
        // even when the captured Win32 detail contains words such as "failed".
        if (eventName.Equals("IMOD.DRIVER.PRELOAD", StringComparison.OrdinalIgnoreCase))
        {
            return DiagnosticLevel.Info;
        }

        if (eventParts.Any(part => part.Equals("FATAL", StringComparison.OrdinalIgnoreCase)))
        {
            return DiagnosticLevel.Fatal;
        }

        if (eventParts.Any(part => part is "ERROR" or "FAILED" or "FAIL"))
        {
            return DiagnosticLevel.Error;
        }

        if (eventName.EndsWith(".ISSUE", StringComparison.OrdinalIgnoreCase)
            || eventParts.Any(part => part is "WARN" or "WARNING" or "MISMATCH" or "CONFLICT" or "TIMEOUT" or "BLOCKED")
            || WarningMessageRegex().IsMatch(message))
        {
            return DiagnosticLevel.Warning;
        }

        if (FailureMessageRegex().IsMatch(message))
        {
            return DiagnosticLevel.Error;
        }

        return DiagnosticLevel.Info;
    }

    private static string InferCategory(string message)
    {
        string eventName = GetEventName(message);
        int dot = eventName.IndexOf('.');
        string category = dot >= 0 ? eventName[..dot] : eventName;
        category = new string(category.Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_').ToArray());
        return string.IsNullOrWhiteSpace(category) ? "GENERAL" : category;
    }

    private static string GetEventName(string message)
    {
        int colon = message.IndexOf(':');
        return (colon >= 0 ? message[..colon] : message).Trim().ToUpperInvariant();
    }

    private static void CloseWriterUnsafe()
    {
        try
        {
            _sessionWriter?.Flush();
            _sessionWriter?.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"DEVICE TWEAKER log close failed: {ex}");
        }
        finally
        {
            _sessionWriter = null;
        }
    }

    private static string Flatten(string value)
    {
        return value
            .Replace("\r\n", " | ", StringComparison.Ordinal)
            .Replace("\n", " | ", StringComparison.Ordinal)
            .Replace("\r", " | ", StringComparison.Ordinal)
            .Trim();
    }

    private static string GetAppVersion()
    {
        try
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string product = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                ?? "unknown";
            string file = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version
                ?? "unknown";
            string assemblyVersion = assembly.GetName().Version?.ToString()
                ?? "unknown";
            return $"product={product} file={file} assembly={assemblyVersion}";
        }
        catch
        {
            return "unknown";
        }
    }

    [GeneratedRegex("""\b(?:failed|exception)(?=[:\s]|$)|\bstatus=(?:FAIL|FAILED)\b|\b(?:errors|failures)=[1-9]\d*\b|\bseverity=(?:Error|Fatal)\b|\berror=(?!""|0\b)(?:"[^"]+"|\S+)""", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex FailureMessageRegex();

    [GeneratedRegex(@"\bseverity=Warning\b|\bwarnings=[1-9]\d*\b|\bstatus=blocked\b|\bblocked=(?:true|[1-9]\d*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex WarningMessageRegex();

    [GeneratedRegex(@"\bstatus=PASS\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex TestPassMessageRegex();
}
