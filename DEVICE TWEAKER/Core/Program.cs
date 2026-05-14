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
        Application.Run(new MainForm());
    }
}
