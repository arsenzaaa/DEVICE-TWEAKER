using System.Diagnostics.CodeAnalysis;
namespace DeviceTweakerCS;

static class Program
{
    [STAThread]
    static void Main()
    {
        if (!WindowsSecurity.IsAdministrator())
        {
            if (WindowsSecurity.TryRelaunchAsAdministrator())
            {
                return;
            }

            MessageBox.Show(
                "This tool must be run as Administrator (it writes to HKLM registry).\n\nRight-click the EXE and choose 'Run as administrator'.",
                "DEVICE TWEAKER",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) => LogFatal("UI", e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogFatal("Domain", ex);
            }
        };

        Application.Run(new MainForm());
    }

    private static void LogFatal(string source, Exception ex)
    {
        AppDiagnostics.WriteFatal(source, ex);
    }
}
