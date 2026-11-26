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

    internal static string BuildXml(string scriptPath, bool enabled = true)
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
                    new XElement(Ns + "Enabled", enabled ? "true" : "false"),
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

    internal static bool IsExplicitlyDisabled(string? registeredXml)
    {
        if (registeredXml is null) return false;
        try
        {
            return IsDisabled(XDocument.Parse(registeredXml).Root?.Element(Ns + "Settings"));
        }
        catch (System.Xml.XmlException) { return false; }
    }

    private static bool IsDisabled(XElement? element)
        => element?.Element(Ns + "Enabled")?.Value.Trim() is "false" or "0";

    internal static string? VerifyRegisteredXml(string registeredXml, string scriptPath, bool enabled = true)
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

        XElement? settings = document.Root?.Element(Ns + "Settings");
        if (IsDisabled(settings) == enabled)
            return enabled ? "the registered task is disabled" : "the registered task should remain disabled";

        XElement[] principals = document.Root?.Element(Ns + "Principals")?.Elements(Ns + "Principal").ToArray() ?? [];
        string? user = principals.Length == 1 ? principals[0].Element(Ns + "UserId")?.Value.Trim() : null;
        if (!string.Equals(user, "S-1-5-18", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(user, @"NT AUTHORITY\SYSTEM", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(user, "SYSTEM", StringComparison.OrdinalIgnoreCase))
            return "the registered task does not run as SYSTEM";
        string? context = (string?)actionLists[0].Attribute("Context");
        if (context is not null && context != (string?)principals[0].Attribute("id"))
            return "the registered task action has a different principal";

        XElement[] triggers = document.Root?.Element(Ns + "Triggers")?.Elements().ToArray() ?? [];
        if (triggers.Length != 3 || triggers.Any(IsDisabled)
            || triggers.Count(t => t.Name == Ns + "LogonTrigger"
                && string.IsNullOrWhiteSpace(t.Element(Ns + "UserId")?.Value)) != 1)
            return "the registered task must have enabled triggers for sign-in, resume and USB/HID arrival";

        XElement expected = XDocument.Parse(BuildXml(scriptPath)).Root!.Element(Ns + "Triggers")!;
        foreach (XElement trigger in expected.Elements(Ns + "EventTrigger"))
        {
            string subscription = trigger.Element(Ns + "Subscription")!.Value;
            XElement? values = trigger.Element(Ns + "ValueQueries");
            if (triggers.Count(t => t.Name == Ns + "EventTrigger"
                && SameSubscription(t.Element(Ns + "Subscription")?.Value, subscription)
                && XNode.DeepEquals(t.Element(Ns + "ValueQueries"), values)) != 1)
                return "a required event trigger or its device query is missing";
        }

        return null;
    }

    private static bool SameSubscription(string? actual, string expected)
    {
        if (actual is null) return false;
        try { return XNode.DeepEquals(XElement.Parse(actual), XElement.Parse(expected)); }
        catch (System.Xml.XmlException) { return false; }
    }

    internal static string? ReadXml()
    {
        object service = Activator.CreateInstance(Type.GetTypeFromProgID("Schedule.Service", throwOnError: true)!)!;
        object? folder = null, task = null;
        try
        {
            ((dynamic)service).Connect();
            folder = ((dynamic)service).GetFolder(@"\");
            try { task = ((dynamic)folder).GetTask(TaskName); }
            catch (Exception ex) when (ex.HResult is unchecked((int)0x80070002) or unchecked((int)0x80070003))
            {
                return null;
            }

            return (string)((dynamic)task).Xml;
        }
        finally
        {
            if (task is not null) System.Runtime.InteropServices.Marshal.ReleaseComObject(task);
            if (folder is not null) System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(service);
        }
    }

    internal static bool Register(string scriptPath, bool enabled, Action<string> log, out string? error)
    {
        if (!WriteRegistration(scriptPath, BuildXml(scriptPath, enabled), out error)) return false;
        string? registered = ReadXml();
        error = registered is null ? "the registered task was not found" : VerifyRegisteredXml(registered, scriptPath, enabled);
        if (error is not null) return false;
        log($"IMOD.STARTUP.TASK: registered {TaskName} script=\"{scriptPath}\"");
        return true;
    }

    internal static bool Restore(string scriptPath, string? previousXml, Action<string> log, out string? error)
    {
        return previousXml is null
            ? Delete(log, out error) : WriteRegistration(scriptPath, previousXml, out error);
    }

    private static bool WriteRegistration(string scriptPath, string xml, out string? error)
    {
        QaRun.DemandWrite("StartupTask.Register");
        // Stage the XML in the protected folder before schtasks reads it.
        string xmlPath = Path.Combine(Path.GetDirectoryName(scriptPath)!, $".task-{Guid.NewGuid():N}.xml");
        try
        {
            AtomicFile.WriteAllText(xmlPath, xml, Encoding.Unicode);
            if (!RunSchtasks(["/Create", "/TN", TaskName, "/XML", xmlPath, "/F"], out string output))
            {
                error = $"schtasks /Create failed: {output}";
                return false;
            }

            error = null;
            return true;
        }
        finally
        {
            try { File.Delete(xmlPath); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
    }

    internal static bool Delete(Action<string> log, out string? error)
    {
        QaRun.DemandWrite("StartupTask.Delete");
        try
        {
            error = null;
            if (ReadXml() is null) return true;

            if (!RunSchtasks(["/Delete", "/TN", TaskName, "/F"], out string output))
            {
                error = $"schtasks /Delete failed: {output}";
                return false;
            }
            if (ReadXml() is not null)
            {
                error = "the startup task is still registered after deletion";
                return false;
            }

            log($"IMOD.STARTUP.TASK: deleted {TaskName}");
            return true;
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException or IOException
            or UnauthorizedAccessException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            error = $"could not confirm startup task deletion: {ex.Message}";
            return false;
        }
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
