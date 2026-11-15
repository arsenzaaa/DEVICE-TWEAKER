using System.Reflection;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using static DeviceTweakerCS.ImodDriverNative;

namespace DeviceTweakerCS;

// Embedded driver/loader resources and their verified files on disk. No GUI state.
internal static class ImodDriverFiles
{
    internal static bool EnsureOnDisk(bool persistDriver, Action<string> log, out string driverPath, out string? error)
    {
        if (QaRun.TryBlock("PUBLISHED DRIVER STAGING"))
        { driverPath = string.Empty; error = "QA read-only session: driver access skipped"; return false; }
        error = null;
        driverPath = ImodDriverPaths.SystemPath();

        try
        {
            using Stream? resource = OpenImodDriverResourceStream();
            if (resource is null)
            {
                error = "embedded DTIMOD.sys not found";
                return false;
            }

            using MemoryStream buffer = new();
            resource.CopyTo(buffer);
            byte[] driverBytes = buffer.ToArray();
            string embeddedHash = ComputeSha256(driverBytes);
            string? expectedHash = ReadEmbeddedImodDriverHash();

            if (!string.IsNullOrWhiteSpace(expectedHash)
                && !HashEquals(embeddedHash, expectedHash))
            {
                error = $"embedded DTIMOD.sys hash mismatch: actual={embeddedHash}, expected={expectedHash}";
                return false;
            }

            string? driverDir = Path.GetDirectoryName(driverPath);
            if (string.IsNullOrWhiteSpace(driverDir))
            {
                error = "failed to resolve IMOD driver directory";
                return false;
            }

            if (File.Exists(driverPath))
            {
                string existingHash = ComputeFileSha256(driverPath);
                if (HashEquals(existingHash, expectedHash ?? embeddedHash))
                {
                    log($"IMOD.DRIVER: using existing staged {driverPath} sha256={existingHash}");
                    return true;
                }

                if (!persistDriver)
                {
                    log(
                        $"IMOD.DRIVER: hash mismatch persistDriver=false -> keep existing {driverPath} " +
                        $"sha256={existingHash} (no replace on refresh/read)");
                    return true;
                }

                log($"IMOD.DRIVER: replacing staged {driverPath} sha256={existingHash} -> {embeddedHash}");
            }
            else if (!persistDriver)
            {
                error = "DTIMOD.sys is not staged; press CHECK to load the driver";
                log($"IMOD.DRIVER: missing {driverPath} persistDriver=false -> skip stage");
                return false;
            }

            if (!ImodDriverPaths.IsSystemPath(driverPath))
            {
                string? driverRoot = Path.GetDirectoryName(driverDir);
                if (string.IsNullOrWhiteSpace(driverRoot))
                {
                    error = "failed to resolve IMOD driver root directory";
                    return false;
                }

                if (!TrySecureImodDriverDirectory(driverRoot, out error))
                {
                    return false;
                }

                if (!TrySecureImodDriverDirectory(driverDir, out error))
                {
                    return false;
                }
            }
            else if (!Directory.Exists(driverDir))
            {
                error = "Windows directory for IMOD driver not found";
                return false;
            }

            AtomicFile.WriteAllBytes(driverPath, driverBytes);

            string writtenHash = ComputeFileSha256(driverPath);
            if (!HashEquals(writtenHash, expectedHash ?? embeddedHash))
            {
                error = $"written DTIMOD.sys hash mismatch: actual={writtenHash}, expected={expectedHash ?? embeddedHash}";
                return false;
            }

            log($"IMOD.DRIVER: staged {driverPath} sha256={writtenHash}");
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to stage DTIMOD.sys: {ex.Message}";
            return false;
        }
    }

    internal static bool TrySecureImodDriverDirectory(string directoryPath, out string? error)
    {
        QaRun.DemandWrite("TrySecureImodDriverDirectory");
        error = null;

        try
        {
            Directory.CreateDirectory(directoryPath);

            DirectorySecurity security = new();
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

            void AddRule(WellKnownSidType sidType)
            {
                security.AddAccessRule(new FileSystemAccessRule(
                    new SecurityIdentifier(sidType, null),
                    FileSystemRights.FullControl,
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags.None,
                    AccessControlType.Allow));
            }

            AddRule(WellKnownSidType.LocalSystemSid);
            AddRule(WellKnownSidType.BuiltinAdministratorsSid);

            new DirectoryInfo(directoryPath).SetAccessControl(security);
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to secure IMOD temp directory {directoryPath}: {ex.Message}";
            return false;
        }
    }

    internal static Stream? OpenImodDriverResourceStream()
    {
        Assembly asm = typeof(ImodDriverFiles).Assembly;
        Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys");
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(".DTIMOD.sys", StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    internal static string? ReadEmbeddedImodDriverHash()
    {
        using Stream? stream = OpenImodDriverHashResourceStream();
        if (stream is null)
        {
            return null;
        }

        using StreamReader reader = new(stream, Encoding.ASCII, detectEncodingFromByteOrderMarks: true);
        return NormalizeHash(reader.ReadToEnd());
    }

    internal static Stream? OpenImodDriverHashResourceStream()
    {
        Assembly asm = typeof(ImodDriverFiles).Assembly;
        Stream? stream = asm.GetManifestResourceStream("DeviceTweakerCS.IMOD.DTIMOD.sys.sha256");
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(".DTIMOD.sys.sha256", StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    internal static string ComputeFileSha256(string path)
    {
        using FileStream stream = File.OpenRead(path);
        return NormalizeHash(Convert.ToHexString(SHA256.HashData(stream))) ?? string.Empty;
    }

    internal static string ComputeSha256(byte[] data)
    {
        return NormalizeHash(Convert.ToHexString(SHA256.HashData(data))) ?? string.Empty;
    }

    internal static string? NormalizeHash(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        StringBuilder sb = new(value.Length);
        foreach (char ch in value)
        {
            if (Uri.IsHexDigit(ch))
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.Length == 0 ? null : sb.ToString();
    }

    internal static bool HashEquals(string? left, string? right)
    {
        string? normalizedLeft = NormalizeHash(left);
        string? normalizedRight = NormalizeHash(right);
        return normalizedLeft is not null
            && normalizedRight is not null
            && string.Equals(normalizedLeft, normalizedRight, StringComparison.Ordinal);
    }

    internal static bool EnsureImodKduPayloadOnDisk(Action<string>? log, out string kduPath, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: EnsureImodKduPayloadOnDisk");
        kduPath = string.Empty;
        error = null;

        string commonData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (string.IsNullOrWhiteSpace(commonData))
        {
            commonData = Path.GetTempPath();
        }

        string payloadRoot = Path.Combine(commonData, "DEVICE TWEAKER", "IMOD", "Loader");
        string kduTarget = Path.Combine(payloadRoot, ImodKduFileName);
        string dbTarget = Path.Combine(payloadRoot, ImodKduDatabaseFileName);

        if (TrySecureImodDriverDirectory(payloadRoot, out error)
            && TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.kdu.exe", ".kdu.exe", kduTarget, out error)
            && TryWriteEmbeddedImodResource("DeviceTweakerCS.IMOD.Loader.drv64.dll", ".drv64.dll", dbTarget, out error))
        {
            kduPath = kduTarget;
            return true;
        }

        log?.Invoke($"IMOD.DRIVER.KDU: embedded payload unavailable error={CompactLogValue(error)}");
        return false;
    }

    internal static bool TryWriteEmbeddedImodResource(string exactName, string suffix, string targetPath, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryWriteEmbeddedImodResource");
        error = null;
        using Stream? resource = OpenManifestResourceStreamExactOrSuffix(exactName, suffix);
        if (resource is null)
        {
            error = $"embedded resource missing: {exactName}";
            return false;
        }

        try
        {
            string? directory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using MemoryStream buffer = new();
            resource.CopyTo(buffer);
            byte[] bytes = buffer.ToArray();
            string payloadHash = ComputeSha256(bytes);

            if (File.Exists(targetPath))
            {
                string existingHash = ComputeFileSha256(targetPath);
                if (HashEquals(existingHash, payloadHash))
                {
                    return true;
                }
            }

            AtomicFile.WriteAllBytes(targetPath, bytes);
            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to write {targetPath}: {ex.Message}";
            return false;
        }
    }

    internal static bool TryVerifyEmbeddedImodResourceFile(string exactName, string suffix, string targetPath, out string? error)
    {
        QaRun.DemandWrite("PUBLISHED RUNTIME: TryVerifyEmbeddedImodResourceFile");
        error = null;
        using Stream? resource = OpenManifestResourceStreamExactOrSuffix(exactName, suffix);
        if (resource is null)
        {
            error = $"embedded resource missing: {exactName}";
            return false;
        }

        if (!File.Exists(targetPath))
        {
            error = $"staged payload missing: {targetPath}";
            return false;
        }

        try
        {
            string expectedHash = NormalizeHash(Convert.ToHexString(SHA256.HashData(resource))) ?? string.Empty;
            string actualHash = ComputeFileSha256(targetPath);
            if (!HashEquals(actualHash, expectedHash))
            {
                error = $"staged payload hash mismatch: {Path.GetFileName(targetPath)}";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = $"failed to verify {targetPath}: {ex.Message}";
            return false;
        }
    }

    internal static Stream? OpenManifestResourceStreamExactOrSuffix(string exactName, string suffix)
    {
        Assembly asm = typeof(ImodDriverFiles).Assembly;
        Stream? stream = asm.GetManifestResourceStream(exactName);
        if (stream is not null)
        {
            return stream;
        }

        foreach (string name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                return asm.GetManifestResourceStream(name);
            }
        }

        return null;
    }

    private static string CompactLogValue(string? value)
        => string.IsNullOrWhiteSpace(value) ? "none" : value.Replace('\r', ' ').Replace('\n', ' ').Trim();
}
