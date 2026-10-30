using System.Reflection;
using System.Security.Cryptography;

namespace DeviceTweakerCS;

internal static class AppBuildIdentity
{
    // MVID identifies the managed build; the digest also distinguishes published EXE variants.
    private static readonly Lazy<string> ExecutableHash = new(() => HashExecutable(Environment.ProcessPath));
    internal static string BuildId => typeof(AppBuildIdentity).Assembly.ManifestModule.ModuleVersionId.ToString("N");
    internal static string ExecutableSha256 => ExecutableHash.Value;

    internal static string ProductVersion
        => typeof(AppBuildIdentity).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown";

    internal static string VersionText
    {
        get
        {
            Assembly assembly = typeof(AppBuildIdentity).Assembly;
            string product = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown";
            string file = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version ?? "unknown";
            return $"product={product} file={file} assembly={assembly.GetName().Version} "
                + $"build={BuildId} exeSha256={ExecutableSha256}";
        }
    }

    internal static string HashExecutable(string? path)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(path)) return "unavailable(no-process-path)";
            using FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return Convert.ToHexString(SHA256.HashData(stream));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return $"unavailable({ex.GetType().Name})";
        }
    }
}
