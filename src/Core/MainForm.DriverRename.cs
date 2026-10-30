using System.Text;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string ImodRuntimeSourceToken = "{{IMOD_RUNTIME_SOURCE}}";

    private static string CreateImodScriptTemplate()
    {
        return ImodScriptTemplate;
    }

    private static string ComposeImodScriptTemplate()
    {
        string script = ReadStartupResource("DeviceTweakerCS.Startup.ApplyIMOD.ps1");
        string runtime = ReadStartupResource("DeviceTweakerCS.Startup.DeviceTweakerImodRuntime.cs");
        if (script.IndexOf(ImodRuntimeSourceToken, StringComparison.Ordinal) < 0)
        {
            throw new InvalidDataException("Startup script template has no runtime source token.");
        }

        return script.Replace(ImodRuntimeSourceToken, runtime.TrimEnd('\n'), StringComparison.Ordinal);
    }

    private static string ReadStartupResource(string name)
    {
        using Stream stream = typeof(MainForm).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidDataException("Missing embedded startup resource: " + name);
        using StreamReader reader = new(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd().Replace("\r\n", "\n", StringComparison.Ordinal);
    }

    private void UpgradeManagedImodStartupForRename()
    {
        DemandRealWrite("UpgradeManagedImodStartupForRename");
        foreach (string path in new[] { GetImodStartupPath(), GetLegacyImodStartupPath() }.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!File.Exists(path)) continue;
            string original = File.ReadAllText(path, Encoding.UTF8);
            string upgraded = UpgradeImodScriptEngine(original);
            string previous = path + ".pre-DTIMOD-" + Guid.NewGuid().ToString("N") + ".bak";
            if (ImodStartupUpgrade.UpdateManaged(original, upgraded, PrepareImodStartupPayload,
                () => File.Copy(path, previous, overwrite: false),
                text => WriteAllTextAtomic(path, text, new UTF8Encoding(false))))
                WriteLog($"IMOD.DRIVER.RENAME: startup upgraded; saved IMOD/NIC ITR values retained; previous={previous}");
        }
    }

    private static string UpgradeImodScriptEngine(string text) => ImodStartupUpgrade.Upgrade(text,
        CreateImodScriptTemplate(), ReadEmbeddedImodDriverHash() ?? string.Empty,
        ComputeEmbeddedImodResourceHash("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe"),
        ComputeEmbeddedImodResourceHash("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll"),
        GetImodDriverSystemPathForScript(), Path.Combine(GetImodStartupKduDirectory(), ImodKduFileName),
        Path.Combine(GetImodStartupKduDirectory(), ImodKduDatabaseFileName));

    private static void PrepareImodStartupPayload()
    {
        if (!EnsureImodStartupKduPayload(out _, out _, out string? error))
            throw new IOException("Could not prepare the saved IMOD loader: " + error);
    }
}
