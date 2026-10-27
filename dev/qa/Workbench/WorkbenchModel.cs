using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceTweakerQa;

internal sealed class DevScenario
{
    public int SchemaVersion { get; set; } = 1;
    public string Id { get; set; } = "";
    public string Preset { get; set; } = "Intel14900K";
    public string Language { get; set; } = "EN";
    public string Filter { get; set; } = "USB";
    public int Width { get; set; } = 1200;
    public int Height { get; set; } = 900;
    public int LayoutDpi { get; set; } = 96;
    public List<DevStep> Steps { get; set; } = [];
    public void Validate()
    {
        DevFiles.Id(Id);
        if (SchemaVersion != 1 || Preset is not ("Intel14900K" or "Ryzen9950X3D" or "Imod" or "NicItr" or "UsbRoles" or "DevicePalette")
            || Language is not ("RU" or "EN") || LayoutDpi is not (96 or 120 or 144) || Width is < 700 or > 2400 || Height is < 500 or > 2000
            || Steps.Count is < 1 or > 200) throw new InvalidDataException("Invalid workbench scenario configuration");
        HashSet<string> ids = [];
        foreach (DevStep step in Steps)
        {
            DevFiles.Id(step.Id);
            if (!ids.Add(step.Id)) throw new InvalidDataException("Duplicate step ID: " + step.Id);
            if (!DevStep.Actions.Contains(step.Action) || step.DurationMs is < 0 or > 10000 || step.IntervalMs is < 50 or > 1000)
                throw new InvalidDataException("Invalid step: " + step.Id);
            if (step.Action == "transition" && (step.Trigger is null || !DevStep.Actions.Contains(step.Trigger.Action)
                || step.Trigger.Action is "transition" or "capture" or "save-state" or "restore-state"))
                throw new InvalidDataException("Transition requires one supported trigger");
        }
    }
}

internal sealed class DevStep
{
    public static readonly HashSet<string> Actions = ["capture", "click", "hover", "leave", "focus", "select", "text", "toggle",
        "assert", "wait", "key", "language", "resize", "scroll", "reveal", "transition", "save-state", "restore-state", "apply-model", "model-error", "gallery", "audit-editors", "dialog", "audit-dialog", "irq-check-model"];
    public string Id { get; set; } = "";
    public string Action { get; set; } = "";
    public string Target { get; set; } = "@main";
    public string Device { get; set; } = "";
    public string Value { get; set; } = "";
    public string Property { get; set; } = "text";
    public int DurationMs { get; set; } = 800;
    public int IntervalMs { get; set; } = 100;
    public int Width { get; set; }
    public int Height { get; set; }
    public bool RequireChange { get; set; } = true;
    public bool RequireSettled { get; set; } = true;
    public DevStep? Trigger { get; set; }
}

internal sealed record DevCapture(string Id, string Path, string Sha256, int Width, int Height, double TimeMs,
    string Tree, string TreeSha256, bool Timeline = false, string Method = "PrintWindow.Client");
internal sealed record DevStateFile(string Path, string Sha256);
internal sealed record DevExecution(string SessionSha256, string ProcessReport, int ProcessId, int ExitCode);
internal sealed record DevStepResult(string Id, string Action, string Status, double TimeMs, string Detail);
internal sealed class DevRun
{
    public int SchemaVersion { get; set; } = 1;
    public string Status { get; set; } = "RUNNING";
    public string ScenarioId { get; set; } = "";
    public string ScenarioSha256 { get; set; } = "";
    public string Environment { get; set; } = "";
    public string Desktop { get; set; } = "";
    public bool ReadOnly { get; set; }
    public List<DevStepResult> Steps { get; set; } = [];
    public List<DevCapture> Captures { get; set; } = [];
    public List<string> Exceptions { get; set; } = [];
    public List<DevStateFile> StateFiles { get; set; } = [];
    public string HardwareValidation { get; set; } = "NOT_PERFORMED";
    public string InputMethod { get; set; } = "Owned WinForms controls / handlers; no physical input";
}

internal static class DevFiles
{
    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };
    public static T Read<T>(string path) => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Json)
        ?? throw new InvalidDataException("Empty JSON: " + path);
    public static void Write<T>(string path, T value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        string temp = path + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(value, Json));
        for (int attempt = 0; ; attempt++)
        {
            try { File.Move(temp, path, overwrite: true); break; }
            catch (Exception ex) when (attempt < 20 && ex is IOException or UnauthorizedAccessException)
            { Thread.Sleep(5); }
        }
    }
    public static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    public static void Id(string value)
    {
        if (string.IsNullOrEmpty(value) || value.Length > 80 || value.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('-' or '_')))
            throw new InvalidDataException("ID must contain ASCII letters, digits, '-' or '_'");
    }
    public static string Inside(string root, string relative)
    {
        if (Path.IsPathRooted(relative)) throw new InvalidDataException("Artifact path must be relative");
        string full = Path.GetFullPath(Path.Combine(root, relative));
        string prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Artifact escapes run directory");
        return full;
    }
    public static void Verify(string root, DevRun run)
    {
        if (run.SchemaVersion != 1 || run.Status != "PASS" || !run.ReadOnly || run.Captures.Count == 0
            || run.Steps.Count == 0 || run.Steps.Any(s => s.Status != "PASS")) throw new InvalidDataException("Incomplete / failed development session");
        if (run.Captures.Select(c => c.Id).Distinct().Count() != run.Captures.Count) throw new InvalidDataException("Duplicate capture ID");
        string scenarioPath = Inside(root, "scenario.json");
        if (Hash(scenarioPath) != run.ScenarioSha256) throw new InvalidDataException("Scenario hash differs from executed input");
        DevScenario scenario = Read<DevScenario>(scenarioPath); scenario.Validate();
        if (scenario.Id != run.ScenarioId || !scenario.Steps.Select(s => (s.Id, s.Action)).SequenceEqual(run.Steps.Select(s => (s.Id, s.Action))))
            throw new InvalidDataException("Executed step sequence differs from scenario");
        foreach (DevStep step in scenario.Steps)
        {
            if (step.Action == "capture" && !run.Captures.Any(c => c.Id == step.Id && !c.Timeline)) throw new InvalidDataException("Missing checkpoint: " + step.Id);
            if (step.Action == "transition" && (run.Captures.Count(c => c.Timeline && c.Id.StartsWith(step.Id + "-frame-", StringComparison.Ordinal)) < 3
                || !run.Captures.Any(c => c.Id == step.Id + "-settled" && !c.Timeline))) throw new InvalidDataException("Missing transition evidence");
        }
        foreach (DevCapture capture in run.Captures)
        {
            Id(capture.Id);
            if (Hash(Inside(root, capture.Path)) != capture.Sha256 || Hash(Inside(root, capture.Tree)) != capture.TreeSha256)
                throw new InvalidDataException("Missing / altered capture: " + capture.Id);
            using System.Drawing.Bitmap image = new(Inside(root, capture.Path));
            if (image.Width != capture.Width || image.Height != capture.Height || capture.Method != "PrintWindow.Client")
                throw new InvalidDataException("Capture dimensions / method do not match evidence");
        }
        foreach (DevStateFile state in run.StateFiles)
            if (Hash(Inside(root, state.Path)) != state.Sha256) throw new InvalidDataException("Altered sandbox state");
    }
    public static void VerifyExecution(string root)
    {
        DevExecution execution = Read<DevExecution>(Inside(root, "execution.json"));
        if (execution.ExitCode != 0 || execution.ProcessId <= 0 || execution.ProcessReport != $"process-{execution.ProcessId}.json"
            || execution.SessionSha256 != Hash(Inside(root, "session.json"))) throw new InvalidDataException("Execution evidence is missing / inconsistent");
        using JsonDocument report = JsonDocument.Parse(File.ReadAllText(Inside(root, execution.ProcessReport)));
        JsonElement data = report.RootElement;
        if (!data.GetProperty("readOnly").GetBoolean() || !data.GetProperty("build").GetProperty("hasTestPanel").GetBoolean()
            || !data.GetProperty("desktop").GetProperty("isolated").GetBoolean()) throw new InvalidDataException("Execution was not a private read-only developer process");
        if (Directory.EnumerateFiles(root, "crash_*.txt", SearchOption.AllDirectories).Any()) throw new InvalidDataException("Crash evidence found");
    }
}
