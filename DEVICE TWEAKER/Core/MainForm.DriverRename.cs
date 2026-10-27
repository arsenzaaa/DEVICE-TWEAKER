using System.Text;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private static string CreateImodScriptTemplate()
    {
        return ImodScriptTemplate;
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
