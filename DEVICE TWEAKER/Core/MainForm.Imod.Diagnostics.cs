using Microsoft.Win32;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const int ImodSystemCodeIntegrityInformationClass = 103;

    private const uint ImodCiOptionEnabled = 0x00000001;
    private const uint ImodCiOptionTestSign = 0x00000002;
    private const uint ImodCiOptionUmciEnabled = 0x00000004;
    private const uint ImodCiOptionTestBuild = 0x00000020;
    private const uint ImodCiOptionHvciEnabled = 0x00000400;
    private const uint ImodCiOptionHvciAudit = 0x00000800;
    private const uint ImodCiOptionHvciStrict = 0x00001000;
    private const uint ImodCiOptionHvciIum = 0x00002000;

    private const string ImodHvciRegistryPath =
        @"SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity";
    private const string ImodDeviceGuardRegistryPath =
        @"SYSTEM\CurrentControlSet\Control\DeviceGuard";
    private const string ImodDeviceGuardPolicyPath =
        @"SOFTWARE\Policies\Microsoft\Windows\DeviceGuard";

    private void LogImodDriverLoadDiagnostics(string driverPath, string? loadError)
    {
        WriteLog($"IMOD.DRIVER.LOAD: failed path={driverPath} error={CompactImodLogValue(loadError)}");
        LogImodDriverSignatureDiagnostics(driverPath);
        LogImodCodeIntegrityDiagnostics(loadError);
    }

    private void LogImodDriverSignatureDiagnostics(string driverPath)
    {
        if (string.IsNullOrWhiteSpace(driverPath) || !File.Exists(driverPath))
        {
            WriteLog($"IMOD.DRIVER.SIGNATURE: file missing path={driverPath}");
            return;
        }

        try
        {
            using X509Certificate baseCertificate = X509Certificate.CreateFromSignedFile(driverPath);
            using X509Certificate2 certificate = new(baseCertificate);
            using X509Chain chain = new();

            chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
            chain.ChainPolicy.VerificationFlags =
                X509VerificationFlags.IgnoreEndRevocationUnknown
                | X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown
                | X509VerificationFlags.IgnoreRootRevocationUnknown;

            bool chainValid = chain.Build(certificate);
            string chainStatus = FormatImodChainStatus(chain);
            string storePresence = FormatImodCertificateStorePresence(certificate);
            string thumbprint = certificate.Thumbprint ?? "none";

            WriteLog(
                "IMOD.DRIVER.SIGNATURE: "
                + $"subject=\"{certificate.Subject}\" "
                + $"thumbprint={thumbprint} "
                + $"chainValid={chainValid} "
                + $"chain={chainStatus} "
                + $"stores={storePresence}");
        }
        catch (Exception ex)
        {
            WriteLog($"IMOD.DRIVER.SIGNATURE: unavailable ({ex.GetType().Name}: {CompactImodLogValue(ex.Message)})");
        }
    }

    private void LogImodCodeIntegrityDiagnostics(string? loadError)
    {
        bool hasCiOptions = TryQueryImodCodeIntegrityOptions(out uint options);
        bool testSigning = hasCiOptions && (options & ImodCiOptionTestSign) != 0;
        bool testBuild = hasCiOptions && (options & ImodCiOptionTestBuild) != 0;
        bool hvciRuntime = hasCiOptions
            && (options & (ImodCiOptionHvciEnabled | ImodCiOptionHvciAudit | ImodCiOptionHvciStrict | ImodCiOptionHvciIum)) != 0;

        if (hasCiOptions)
        {
            WriteLog(
                "IMOD.CI: "
                + $"options=0x{options:X} "
                + $"enabled={(options & ImodCiOptionEnabled) != 0} "
                + $"umci={(options & ImodCiOptionUmciEnabled) != 0} "
                + $"testSign={testSigning} "
                + $"testBuild={testBuild} "
                + $"hvciRuntime={hvciRuntime}");
        }
        else
        {
            WriteLog("IMOD.CI: options=unavailable");
        }

        string hvciRegistry = ReadImodRegistryDwordState(
            RegistryHive.LocalMachine,
            ImodHvciRegistryPath,
            "Enabled");
        string hvciPolicy = ReadImodRegistryDwordState(
            RegistryHive.LocalMachine,
            ImodDeviceGuardPolicyPath,
            "HypervisorEnforcedCodeIntegrity");
        string vbsRegistry = ReadImodRegistryDwordState(
            RegistryHive.LocalMachine,
            ImodDeviceGuardRegistryPath,
            "EnableVirtualizationBasedSecurity");

        WriteLog(
            "IMOD.CI.REGISTRY: "
            + $"hvci={hvciRegistry} "
            + $"hvciPolicy={hvciPolicy} "
            + $"vbs={vbsRegistry}");

        bool signatureRejected = IsImodSignatureRejectedError(loadError);
        if (signatureRejected && hasCiOptions && !testSigning && !testBuild)
        {
            WriteLog("IMOD.CI.HINT: test-signing is off; Windows can reject self-signed or test-signed kernel drivers even when Authenticode trust is valid.");
            WriteLog("IMOD.CI.HINT: enable Windows test mode for this development driver with: bcdedit /set testsigning on; then reboot. Disable it later with: bcdedit /set testsigning off; then reboot.");
        }

        if (signatureRejected && (hvciRuntime || IsEnabledRegistryState(hvciRegistry) || IsEnabledRegistryState(hvciPolicy)))
        {
            WriteLog("IMOD.CI.HINT: HVCI is enabled; unsigned, self-signed, or test-signed kernel drivers are commonly blocked.");
        }
    }

    private static string GetImodKernelCiBlockState()
    {
        if (!TryQueryImodCodeIntegrityOptions(out uint options))
        {
            return "unavailable";
        }

        bool testSigning = (options & ImodCiOptionTestSign) != 0;
        bool testBuild = (options & ImodCiOptionTestBuild) != 0;
        bool hvciRuntime = (options & (ImodCiOptionHvciEnabled | ImodCiOptionHvciAudit | ImodCiOptionHvciStrict | ImodCiOptionHvciIum)) != 0;
        return $"enabled={(options & ImodCiOptionEnabled) != 0} testSign={testSigning} testBuild={testBuild} hvciRuntime={hvciRuntime}";
    }

    private static bool IsImodKernelCiBlockedLoadError(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return false;
        }

        return error.Contains("577", StringComparison.OrdinalIgnoreCase)
            || error.Contains("digital signature", StringComparison.OrdinalIgnoreCase)
            || error.Contains("цифров", StringComparison.OrdinalIgnoreCase)
            || error.Contains("signature", StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatImodChainStatus(X509Chain chain)
    {
        if (chain.ChainStatus.Length == 0)
        {
            return "ok";
        }

        return string.Join("|", chain.ChainStatus.Select(s => s.Status.ToString()));
    }

    private static string FormatImodCertificateStorePresence(X509Certificate2 certificate)
    {
        List<string> entries = [];
        AppendImodStorePresence(entries, certificate, StoreLocation.CurrentUser, StoreName.Root, "CU-Root");
        AppendImodStorePresence(entries, certificate, StoreLocation.CurrentUser, StoreName.TrustedPublisher, "CU-TrustedPublisher");
        AppendImodStorePresence(entries, certificate, StoreLocation.LocalMachine, StoreName.Root, "LM-Root");
        AppendImodStorePresence(entries, certificate, StoreLocation.LocalMachine, StoreName.TrustedPublisher, "LM-TrustedPublisher");
        return string.Join(",", entries);
    }

    private static void AppendImodStorePresence(
        List<string> entries,
        X509Certificate2 certificate,
        StoreLocation location,
        StoreName name,
        string label)
    {
        string? thumbprint = certificate.Thumbprint;
        if (string.IsNullOrWhiteSpace(thumbprint))
        {
            entries.Add($"{label}=no-thumbprint");
            return;
        }

        try
        {
            using X509Store store = new(name, location);
            store.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadOnly);
            X509Certificate2Collection matches = store.Certificates.Find(
                X509FindType.FindByThumbprint,
                thumbprint,
                validOnly: false);
            entries.Add($"{label}={(matches.Count > 0 ? "yes" : "no")}");
        }
        catch (Exception ex)
        {
            entries.Add($"{label}=error:{ex.GetType().Name}");
        }
    }

    private static bool TryQueryImodCodeIntegrityOptions(out uint options)
    {
        ImodCodeIntegrityInformation info = new()
        {
            Length = (uint)Marshal.SizeOf<ImodCodeIntegrityInformation>()
        };

        int status = NtQuerySystemInformation(
            ImodSystemCodeIntegrityInformationClass,
            ref info,
            Marshal.SizeOf<ImodCodeIntegrityInformation>(),
            IntPtr.Zero);

        if (status < 0)
        {
            options = 0;
            return false;
        }

        options = info.CodeIntegrityOptions;
        return true;
    }

    private static string ReadImodRegistryDwordState(RegistryHive hive, string subKeyName, string valueName)
    {
        try
        {
            using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
            using RegistryKey? key = baseKey.OpenSubKey(subKeyName);
            object? rawValue = key?.GetValue(valueName);
            if (rawValue is null)
            {
                return "missing";
            }

            uint value = rawValue switch
            {
                int intValue => unchecked((uint)intValue),
                long longValue => unchecked((uint)longValue),
                byte[] bytes when bytes.Length >= sizeof(uint) => BitConverter.ToUInt32(bytes, 0),
                _ => Convert.ToUInt32(rawValue, CultureInfo.InvariantCulture)
            };

            return $"{value}({(value == 0 ? "off" : "on")})";
        }
        catch (Exception ex)
        {
            return $"error:{ex.GetType().Name}";
        }
    }

    private static bool IsEnabledRegistryState(string value)
    {
        return value.EndsWith("(on)", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImodSignatureRejectedError(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return false;
        }

        return error.Contains("signature", StringComparison.OrdinalIgnoreCase)
            || error.Contains("digital", StringComparison.OrdinalIgnoreCase)
            || error.Contains("подпис", StringComparison.OrdinalIgnoreCase)
            || error.Contains("цифров", StringComparison.OrdinalIgnoreCase)
            || error.Contains("подпис", StringComparison.OrdinalIgnoreCase)
            || error.Contains("цифров", StringComparison.OrdinalIgnoreCase)
            || error.Contains("577", StringComparison.OrdinalIgnoreCase);
    }

    private static string CompactImodLogValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "none";
        }

        string compact = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return compact.Length <= 240 ? compact : compact[..240] + "...";
    }

    [DllImport("ntdll.dll")]
    private static extern int NtQuerySystemInformation(
        int systemInformationClass,
        ref ImodCodeIntegrityInformation systemInformation,
        int systemInformationLength,
        IntPtr returnLength);

    [StructLayout(LayoutKind.Sequential)]
    private struct ImodCodeIntegrityInformation
    {
        public uint Length;
        public uint CodeIntegrityOptions;
    }
}
