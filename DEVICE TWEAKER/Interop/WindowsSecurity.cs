using System.Diagnostics;
using System.Security.Principal;

namespace DeviceTweakerCS;

internal static class WindowsSecurity
{
    public static bool IsAdministrator()
    {
        try
        {
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    public static bool TryRelaunchAsAdministrator()
    {
        try
        {
            string? exePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exePath))
            {
                return false;
            }

            string[] args = Environment.GetCommandLineArgs();
            string arguments = args.Length > 1 ? string.Join(" ", args.Skip(1).Select(QuoteArgument)) : string.Empty;

            ProcessStartInfo startInfo = new()
            {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = true,
                Verb = "runas",
            };

            Process.Start(startInfo);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string QuoteArgument(string arg)
    {
        if (string.IsNullOrEmpty(arg))
        {
            return "\"\"";
        }

        if (arg.Contains(' ') || arg.Contains('\t') || arg.Contains('"'))
        {
            return "\"" + arg.Replace("\"", "\\\"") + "\"";
        }

        return arg;
    }
}
