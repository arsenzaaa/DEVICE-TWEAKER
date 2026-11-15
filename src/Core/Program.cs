using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace DeviceTweakerCS;

static class Program
{
    private const string SingleInstanceMutexName = @"Global\DEVICE_TWEAKER_SINGLE_INSTANCE";
    private static int _handlingFatalUiException;

    [STAThread]
    static void Main(string[] args)
    {
        if (args.Contains("--qa-read-only", StringComparer.Ordinal))
            Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_READ_ONLY", "1");
        try { QaDesktop.DemandAssignedDesktop(); }
        catch (Exception ex)
        {
            // Refuse before logging setup, single-instance handling or window creation.
            Console.Error.WriteLine(ex.Message);
            Environment.ExitCode = 3;
            return;
        }
        if (QaRun.IsReadOnly)
        {
            AppDomain.CurrentDomain.ProcessExit += (_, _) => QaRun.WriteReport();
            QaRun.WriteReport();
            if (args.Contains("--qa-probe", StringComparer.Ordinal)) return;
        }

        UiLanguage.Initialize();
        // The manifest requires elevation; this only triggers when UAC is off
        // and the account is not an administrator, where relaunching cannot help.
        if (!QaRun.IsReadOnly && !WindowsSecurity.IsAdministrator())
        {
            MessageBox.Show(
                UiLanguage.Text("Run DEVICE TWEAKER as administrator to apply settings.\n\nRight-click the EXE and choose 'Run as administrator'."),
                "DEVICE TWEAKER",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        string mutexName = QaRun.IsReadOnly
            ? @"Local\DEVICE_TWEAKER_QA_" + QaRun.SessionId
            : SingleInstanceMutexName;
        using Mutex singleInstanceMutex = new(initiallyOwned: true, mutexName, out bool isFirstInstance);
        if (!isFirstInstance)
        {
            if (!QaRun.IsReadOnly) ActivateExistingInstance();
            return;
        }

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) => HandleFatalUiException(e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogFatal("Domain", ex);
            }
        };

        Application.Run(new MainForm());
        GC.KeepAlive(singleInstanceMutex);
    }

    private static void ActivateExistingInstance()
    {
        // The first process can still be starting when the second process reaches
        // this point, so briefly wait for its main window to be created.
        for (int attempt = 0; attempt < 30; attempt++)
        {
            IntPtr existingWindow = FindMainWindow();
            if (existingWindow != IntPtr.Zero)
            {
                if (NativeUser32.IsIconic(existingWindow))
                {
                    NativeUser32.ShowWindow(existingWindow, NativeUser32.SwRestore);
                }

                NativeUser32.SetForegroundWindow(existingWindow);
                return;
            }

            Thread.Sleep(100);
        }
    }

    private static IntPtr FindMainWindow()
    {
        IntPtr result = IntPtr.Zero;
        uint currentProcessId = (uint)Environment.ProcessId;
        NativeUser32.EnumWindows((window, _) =>
        {
            if (!NativeUser32.IsWindowVisible(window))
            {
                return true;
            }

            StringBuilder title = new(128);
            NativeUser32.GetWindowText(window, title, title.Capacity);
            StringBuilder className = new(256);
            NativeUser32.GetClassName(window, className, className.Capacity);
            NativeUser32.GetWindowThreadProcessId(window, out uint windowProcessId);
            if (!IsInstanceWindow(title.ToString(), className.ToString(), windowProcessId, currentProcessId))
            {
                return true;
            }

            result = window;
            return false;
        }, IntPtr.Zero);
        return result;
    }

    // An Explorer window opened on a folder named "DEVICE TWEAKER" has the same title,
    // so the window must also be a WinForms window of another process.
    internal static bool IsInstanceWindow(string title, string className, uint windowProcessId, uint currentProcessId)
        => string.Equals(title, "DEVICE TWEAKER", StringComparison.Ordinal)
            && className.StartsWith("WindowsForms10.", StringComparison.Ordinal)
            && windowProcessId != 0
            && windowProcessId != currentProcessId;

    private static void LogFatal(string source, Exception ex)
    {
        AppDiagnostics.WriteFatal(source, ex);
    }

    private static void HandleFatalUiException(Exception exception)
    {
        LogFatal("UI", exception);
        if (Interlocked.Exchange(ref _handlingFatalUiException, 1) != 0)
        {
            Environment.FailFast("Repeated unhandled UI exception.", exception);
        }

        try
        {
            MessageBox.Show(
                UiLanguage.Text("DEVICE TWEAKER has crashed and will close.\n\n"
                    + "A crash report is saved in logs. No further settings will be applied."),
                UiLanguage.Text("DEVICE TWEAKER: CRITICAL ERROR"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            Environment.ExitCode = 1;
            Application.ExitThread();
        }
    }
}
