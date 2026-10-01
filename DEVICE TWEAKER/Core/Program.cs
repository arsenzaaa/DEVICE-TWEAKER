namespace DeviceTweakerCS;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        bool applyImod = HasArg(args, "--apply-imod");
        string? imodScriptPath = GetArgValue(args, "--imod-script");

        if (!WindowsSecurity.IsAdministrator())
        {
            if (WindowsSecurity.TryRelaunchAsAdministrator())
            {
                return;
            }

            if (applyImod)
            {
                Environment.ExitCode = 1;
                return;
            }

            MessageBox.Show(
                "This tool must be run as Administrator (it writes to HKLM registry).\n\nRight-click the EXE and choose 'Run as administrator'.",
                "DEVICE TWEAKER",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        if (applyImod)
        {
            int exitCode = MainForm.ApplyImodFromScript(imodScriptPath, out _);
            Environment.ExitCode = exitCode;
            return;
        }

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }

    private static bool HasArg(string[] args, string name)
    {
        return args.Any(arg => string.Equals(arg, name, StringComparison.OrdinalIgnoreCase));
    }

    private static string? GetArgValue(string[] args, string name)
    {
        string prefix = name + "=";
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return arg[prefix.Length..].Trim('"');
            }

            if (string.Equals(arg, name, StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                return args[i + 1].Trim('"');
            }
        }

        return null;
    }
}
