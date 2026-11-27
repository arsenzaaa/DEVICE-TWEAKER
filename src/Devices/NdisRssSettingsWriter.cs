using Microsoft.Win32;
using System.Globalization;

namespace DeviceTweakerCS;

internal interface INdisRssKey : IRegistryValues, IDisposable { }

// Registry settings only. Active RSS readback belongs to NdisRssReader; a
// successful write does not mean that the adapter has adopted the setting.
internal sealed class NdisRssSettingsWriter
{
    private readonly Func<string, string?> _resolveClassKey;
    private readonly Func<string, bool, INdisRssKey?> _openKey;
    private readonly Action<string> _log;

    internal NdisRssSettingsWriter(Func<string, string?> resolveClassKey, Action<string> log)
        : this(resolveClassKey, OpenWindowsClassKey, log) { }

    // Fake keys exercise the same write sequence without accessing Windows.
    internal NdisRssSettingsWriter(Func<string, string?> resolveClassKey,
        Func<string, bool, INdisRssKey?> openKey, Action<string> log)
    {
        _resolveClassKey = resolveClassKey;
        _openKey = openKey;
        _log = log;
    }

    internal static INdisRssKey? OpenWindowsClassKey(string path, bool create)
    {
        QaRun.DemandWrite("NdisRssSettings.OpenClassKey");
        RegistryKey? key = create ? Registry.LocalMachine.CreateSubKey(path)
            : Registry.LocalMachine.OpenSubKey(path, writable: true);
        return key is null ? null : new WindowsKey(key);
    }

    internal void SetBaseCore(string instance, int value, OperationReport? report, string? name)
        => SetSingle(instance, "*RssBaseProcNumber", Math.Max(0, value), "set RSS base CPU", report, name);

    internal void SetQueues(string instance, int value, OperationReport? report, string? name)
        => SetSingle(instance, "*NumRssQueues", Math.Max(1, value), "set RSS queue count", report, name);

    internal void SetExtra(string instance, NdisRssLayout.Layout layout, OperationReport? report, string? name)
        => Run(instance, create: true, "RSS.EXTRA.SET", "set extended RSS values", report, name,
            values =>
            {
                values.Set("*RssBaseProcGroup", 0, RegistryValueKind.DWord);
                values.Set("*MaxRssProcessors", layout.Processors.Count, RegistryValueKind.DWord);
                values.Set("*RSSMaxProcGroup", 0, RegistryValueKind.DWord);
                values.Set("*RssMaxProcNumber", layout.MaxProcessor, RegistryValueKind.DWord);
                // Explicit NUMAStatic (4) so ndis.sys does not default to Balanced (6)
                // on hybrid CPUs and ignore *RssBaseProcNumber / *RssMaxProcNumber.
                values.Set("*RSSProfile", 4, RegistryValueKind.DWord);
                // Keep NDIS's NUMA selection; do not force node 0.
                values.Delete("*NumaNodeId");
            }, $"group=0 maxProcessors={layout.Processors.Count} maxProc={layout.MaxProcessor} cpus=[{string.Join(',', layout.Processors)}] profile=NUMAStatic(4) numa=automatic");

    internal void ClearBaseCore(string instance, OperationReport? report, string? name)
        => ClearSingle(instance, "*RssBaseProcNumber", "clear RSS base CPU", report, name);

    internal void ClearQueues(string instance, OperationReport? report, string? name)
        => ClearSingle(instance, "*NumRssQueues", "clear RSS queue count", report, name);

    internal void ClearExtra(string instance, OperationReport? report, string? name)
        => Run(instance, create: false, "RSS.CLEAR", "clear extended RSS values", report, name,
            values =>
            {
                foreach (string value in new[] { "*RssBaseProcGroup", "*MaxRssProcessors", "*RSSMaxProcGroup", "*RssMaxProcNumber", "*RSSProfile", "*NumaNodeId" })
                    values.Delete(value);
            }, "values=extra-rss keyPresent=true");

    private void SetSingle(string instance, string valueName, int value, string context,
        OperationReport? report, string? name)
        => Run(instance, create: true, "RSS.SET", context, report, name,
            values => values.Set(valueName, value, RegistryValueKind.DWord),
            $"-> {valueName}={value} (class key)");

    private void ClearSingle(string instance, string valueName, string context, OperationReport? report, string? name)
        => Run(instance, create: false, "RSS.CLEAR", context, report, name,
            values => values.Delete(valueName), $"value={valueName} keyPresent=true");

    private void Run(string instance, bool create, string prefix, string context,
        OperationReport? report, string? name, Action<IRegistryValues> operation, string success)
    {
        string? path = null;
        string stage = "resolve-key";
        LoggedValues? values = null;
        try
        {
            path = _resolveClassKey(instance);
            if (string.IsNullOrWhiteSpace(path))
            {
                _log($"{prefix}.SKIP: {instance} operation=\"{context}\" reason=class-key-not-found status={(create ? "not-written" : "not-present")}");
                if (create) report?.AddError($"{name ?? instance}: {context}", "NDIS class key was not found");
                return;
            }
            stage = "open-key";
            using (INdisRssKey? key = _openKey(path, create))
            {
                if (key is null)
                {
                    if (create) throw new InvalidOperationException("NDIS class key could not be opened for writing");
                    _log($"{prefix}.SKIP: {instance} operation=\"{context}\" path=\"HKLM\\{path}\" reason=class-key-unavailable status=not-present");
                    return;
                }
                values = new LoggedValues(key, instance, path, _log);
                stage = "values";
                operation(values);
                stage = "close-key";
            }
            _log($"{prefix}: {instance} {success}");
        }
        catch (Exception error) when (error is not QaWriteBlockedException)
        {
            int completed = values?.Completed ?? 0;
            _log($"{prefix}.ERROR: {instance} operation=\"{context}\" path=\"HKLM\\{path}\" stage={stage} "
                + $"value={values?.AttemptedValue ?? "none"} action={values?.AttemptedAction ?? "none"} requested={values?.Requested ?? "none"} "
                + $"completed={completed} status={(completed > 0 ? "partial" : "failed")} hresult=0x{error.HResult:X8} error=\"{SingleLine(error.ToString())}\"");
            report?.AddError($"{name ?? instance}: {context}", error.Message);
        }
    }

    private static string SingleLine(string text) => text.Replace("\r\n", " | ")
        .Replace("\n", " | ").Replace("\r", " | ").Replace('"', '\'').Trim();

    private sealed class WindowsKey : INdisRssKey
    {
        private readonly RegistryKey _key;
        private readonly WindowsRegistryValues _values;

        internal WindowsKey(RegistryKey key)
        {
            _key = key;
            _values = new(key);
        }

        public void Set(string name, object value, RegistryValueKind kind) => _values.Set(name, value, kind);
        public void Delete(string name) => _values.Delete(name);
        public void Dispose() => _key.Dispose();
    }

    private sealed class LoggedValues(IRegistryValues key, string instance, string path, Action<string> log) : IRegistryValues
    {
        internal int Completed { get; private set; }
        internal string? AttemptedValue { get; private set; }
        internal string? AttemptedAction { get; private set; }
        internal string? Requested { get; private set; }

        public void Set(string name, object value, RegistryValueKind kind)
        {
            AttemptedValue = name; AttemptedAction = "set";
            Requested = Convert.ToString(value, CultureInfo.InvariantCulture);
            key.Set(name, value, kind);
            Completed++;
            log($"RSS.REG.SET: {instance} path=\"HKLM\\{path}\" value={name} requested={value} kind={kind} status=written");
        }

        public void Delete(string name)
        {
            AttemptedValue = name; AttemptedAction = "delete";
            Requested = "absent";
            key.Delete(name);
            Completed++;
            log($"RSS.REG.CLEAR: {instance} path=\"HKLM\\{path}\" value={name} status=cleared-or-absent");
        }
    }
}
