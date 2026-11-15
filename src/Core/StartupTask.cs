using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace DeviceTweakerCS;

// Scheduled task that runs Startup\ApplyIMOD.ps1 as SYSTEM. It is visible in
// Task Scheduler under "DEVICE TWEAKER" so users can see what runs and when.
internal static class StartupTask
{
    internal const string TaskName = @"\DEVICE TWEAKER\IMOD";
    internal const string PowerShellCommand = @"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe";
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";

    internal static string BuildArguments(string scriptPath)
        => "-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass"
            + $" -File \"{scriptPath}\" -TriggerDevice \"$(Device)\"";

    internal static string BuildXml(string scriptPath)
    {
        string powerShell = PowerShellCommand;
        string arguments = BuildArguments(scriptPath);

        XElement EventTrigger(string delay, string logPath, string selector, bool passDevice)
        {
            string query = $"<QueryList><Query Id=\"0\" Path=\"{logPath}\"><Select Path=\"{logPath}\">{selector}</Select></Query></QueryList>";
            XElement trigger = new(Ns + "EventTrigger",
                new XElement(Ns + "Enabled", "true"),
                new XElement(Ns + "Delay", delay),
                new XElement(Ns + "Subscription", query));
            if (passDevice)
            {
                trigger.Add(new XElement(Ns + "ValueQueries",
                    new XElement(Ns + "Value",
                        new XAttribute("name", "Device"),
                        "Event/EventData/Data[@Name='DeviceInstanceId']")));
            }

            return trigger;
        }

        XDocument document = new(
            new XDeclaration("1.0", "UTF-16", null),
            new XElement(Ns + "Task",
                new XAttribute("version", "1.4"),
                new XElement(Ns + "RegistrationInfo",
                    new XElement(Ns + "Author", "DEVICE TWEAKER"),
                    new XElement(Ns + "Description",
                        $"Applies the saved USB IMOD / NIC ITR profile. Script and log: {Path.GetDirectoryName(scriptPath)}"),
                    new XElement(Ns + "URI", TaskName)),
                new XElement(Ns + "Triggers",
                    new XElement(Ns + "LogonTrigger",
                        new XElement(Ns + "Enabled", "true"),
                        new XElement(Ns + "Delay", "PT5S")),
                    // Resume from sleep / hibernation: drivers reinitialize interrupt moderation.
                    EventTrigger("PT5S", "System",
                        "*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]", passDevice: false),
                    // A started device can receive another xHCI interrupter; the script ignores non-USB / HID devices.
                    EventTrigger("PT3S", "Microsoft-Windows-Kernel-PnP/Configuration",
                        "*[System[EventID=410]]", passDevice: true)),
                new XElement(Ns + "Principals",
                    new XElement(Ns + "Principal",
                        new XAttribute("id", "System"),
                        new XElement(Ns + "UserId", "S-1-5-18"),
                        new XElement(Ns + "RunLevel", "HighestAvailable"))),
                new XElement(Ns + "Settings",
                    new XElement(Ns + "MultipleInstancesPolicy", "IgnoreNew"),
                    new XElement(Ns + "DisallowStartIfOnBatteries", "false"),
                    new XElement(Ns + "StopIfGoingOnBatteries", "false"),
                    new XElement(Ns + "AllowHardTerminate", "true"),
                    new XElement(Ns + "StartWhenAvailable", "false"),
                    new XElement(Ns + "RunOnlyIfNetworkAvailable", "false"),
                    new XElement(Ns + "IdleSettings",
                        new XElement(Ns + "StopOnIdleEnd", "false"),
                        new XElement(Ns + "RestartOnIdle", "false")),
                    new XElement(Ns + "AllowStartOnDemand", "true"),
                    new XElement(Ns + "Enabled", "true"),
                    new XElement(Ns + "Hidden", "false"),
                    new XElement(Ns + "RunOnlyIfIdle", "false"),
                    new XElement(Ns + "ExecutionTimeLimit", "PT2M"),
                    new XElement(Ns + "Priority", "4")),
                new XElement(Ns + "Actions",
                    new XAttribute("Context", "System"),
                    new XElement(Ns + "Exec",
                        new XElement(Ns + "Command", powerShell),
                        new XElement(Ns + "Arguments", arguments)))));

        using StringWriter writer = new Utf16StringWriter();
        document.Save(writer);
        return writer.ToString();
    }

    // Checks the task Windows actually stored: exactly one action, the expected
    // PowerShell command and arguments, nothing else that SYSTEM would run.
    internal static string? VerifyRegisteredXml(string registeredXml, string scriptPath)
    {
        XDocument document;
        try
        {
            document = XDocument.Parse(registeredXml);
        }
        catch (System.Xml.XmlException ex)
        {
            return $"the registered task XML could not be parsed: {ex.Message}";
        }

        List<XElement> actionLists = document.Descendants(Ns + "Actions").ToList();
        if (actionLists.Count != 1)
        {
            return $"the registered task has {actionLists.Count} action lists, expected 1";
        }

        List<XElement> actions = actionLists[0].Elements().ToList();
        if (actions.Count != 1 || actions[0].Name != Ns + "Exec")
        {
            return $"the registered task has {actions.Count} actions ({string.Join(",", actions.Select(a => a.Name.LocalName))}), expected one Exec";
        }

        string? command = actions[0].Element(Ns + "Command")?.Value.Trim();
        string? arguments = actions[0].Element(Ns + "Arguments")?.Value.Trim();
        if (!string.Equals(command, PowerShellCommand, StringComparison.OrdinalIgnoreCase))
        {
            return $"the registered task runs an unexpected command: {command}";
        }

        if (!string.Equals(arguments, BuildArguments(scriptPath), StringComparison.Ordinal))
        {
            return $"the registered task has unexpected arguments: {arguments}";
        }

        return null;
    }

    internal static bool Register(string scriptPath, Action<string> log, out string? error)
    {
        QaRun.DemandWrite("StartupTask.Register");
        // The XML is staged next to the script, in the folder only administrators and SYSTEM
        // can change, so another process cannot edit it before schtasks reads it.
        string? stagingDirectory = Path.GetDirectoryName(Path.GetFullPath(scriptPath));
        if (string.IsNullOrWhiteSpace(stagingDirectory))
        {
            error = $"cannot resolve the startup folder for {scriptPath}";
            return false;
        }

        string xmlPath = Path.Combine(stagingDirectory, $".task-{Guid.NewGuid():N}.xml");
        try
        {
            using (FileStream stream = new(xmlPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (StreamWriter writer = new(stream, Encoding.Unicode))
            {
                writer.Write(BuildXml(scriptPath));
            }

            if (!RunSchtasks(["/Create", "/TN", TaskName, "/XML", xmlPath, "/F"], out string output))
            {
                error = $"schtasks /Create failed: {output}";
                return false;
            }

            if (!RunSchtasks(["/Query", "/TN", TaskName, "/XML"], out string registered))
            {
                error = $"schtasks /Query failed after registration: {registered}";
                return false;
            }

            string? verifyError = VerifyRegisteredXml(registered, scriptPath);
            if (verifyError is not null)
            {
                error = verifyError;
                return false;
            }

            log($"IMOD.STARTUP.TASK: registered {TaskName} script=\"{scriptPath}\"");
            error = null;
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException
            or System.ComponentModel.Win32Exception)
        {
            error = ex.Message;
            return false;
        }
        finally
        {
            try
            {
                File.Delete(xmlPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A leftover task XML in the protected startup folder is harmless.
            }
        }
    }

    internal static bool Exists() => RunSchtasks(["/Query", "/TN", TaskName], out _);

    internal static bool Delete(Action<string> log, out string? error)
    {
        QaRun.DemandWrite("StartupTask.Delete");
        if (!Exists())
        {
            error = null;
            return true;
        }

        if (!RunSchtasks(["/Delete", "/TN", TaskName, "/F"], out string output))
        {
            error = $"schtasks /Delete failed: {output}";
            return false;
        }

        log($"IMOD.STARTUP.TASK: deleted {TaskName}");
        error = null;
        return true;
    }

    private static bool RunSchtasks(IReadOnlyList<string> arguments, out string output)
    {
        ProcessStartInfo info = new(Path.Combine(Environment.SystemDirectory, "schtasks.exe"))
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (string argument in arguments)
        {
            info.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(info) ?? throw new InvalidOperationException("schtasks.exe did not start");
        Task<string> stdout = process.StandardOutput.ReadToEndAsync();
        Task<string> stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30000))
        {
            process.Kill(entireProcessTree: true);
            output = "timeout";
            return false;
        }

        output = (stdout.Result + stderr.Result).Trim();
        return process.ExitCode == 0;
    }

    private sealed class Utf16StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.Unicode;
    }
}
