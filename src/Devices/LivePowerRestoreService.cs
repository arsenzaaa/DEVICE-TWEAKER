namespace DeviceTweakerCS;

internal sealed record LivePowerRestoreTarget(string InstanceId, bool? Enabled);
internal sealed record LivePowerWriteResult(DevicePowerPolicy.WriteStatus Status, string? Error);

// Runs against the production WMI adapter or a model. The target value is read
// from restored persistence, never from the still-unrestored live checkbox.
internal static class LivePowerRestoreService
{
    internal static bool? ReadRestoredEnabled(DeviceTweakerBackup backup, string instanceId, string? classKey, bool usb)
    {
        RegistryValueBackup? Find(string? path, string name) => path is null ? null : backup.RegistryValues.FirstOrDefault(value =>
            string.Equals(value.Hive, "HKLM", StringComparison.OrdinalIgnoreCase)
            && string.Equals(value.Path, path, StringComparison.OrdinalIgnoreCase)
            && string.Equals(value.Name, name, StringComparison.OrdinalIgnoreCase));
        static bool TryDword(RegistryValueBackup value, out int number)
        {
            number = 0;
            return string.Equals(value.Kind, "DWord", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(value.Data, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out number);
        }

        RegistryValueBackup? caps = Find(classKey, DevicePowerPolicy.PnPCapabilitiesName);
        if (string.Equals(caps?.Kind, "ReadError", StringComparison.OrdinalIgnoreCase)) return null;
        if (caps is { Exists: true })
            return TryDword(caps, out int bits) ? DevicePowerPolicy.IsNicTurnOffAllowed(bits) : null;
        if (!usb)
            return caps is { Exists: false } ? true : null;

        string parameters = UsbSelectiveSuspendPolicy.DeviceParametersPath(instanceId);
        RegistryValueBackup? selective = Find(parameters, UsbSelectiveSuspendPolicy.SelectiveSuspendEnabledName);
        RegistryValueBackup? enhanced = Find(parameters, UsbSelectiveSuspendPolicy.EnhancedPowerManagementEnabledName);
        if (selective is null || enhanced is null) return null;
        if (string.Equals(selective.Kind, "ReadError", StringComparison.OrdinalIgnoreCase)
            || string.Equals(enhanced.Kind, "ReadError", StringComparison.OrdinalIgnoreCase)) return null;
        if (selective.Exists && !TryDword(selective, out _)
            || enhanced.Exists && !TryDword(enhanced, out _)) return null;
        return (!selective.Exists || TryDword(selective, out int s) && s != 0)
            && (!enhanced.Exists || TryDword(enhanced, out int e) && e != 0);
    }

    internal static OperationReport Restore(
        IEnumerable<LivePowerRestoreTarget> targets,
        Func<string, bool, LivePowerWriteResult> write,
        Action activatePowerScheme, Action<string> log)
    {
        var report = new OperationReport();
        foreach (var group in targets.GroupBy(target => target.InstanceId, StringComparer.OrdinalIgnoreCase))
        {
            string component = group.Key + ": live power";
            bool?[] values = group.Select(target => target.Enabled).Distinct().ToArray();
            if (values.Length != 1)
            {
                report.AddError(component, "Restored power values disagree.", "No live write was made.");
                continue;
            }
            if (!values[0].HasValue)
            {
                report.AddSkipped(component, "The restored power setting is unavailable.");
                continue;
            }
            try
            {
                LivePowerWriteResult result = write(group.Key, values[0]!.Value);
                log($"BACKUP.RESTORE.POWER: {group.Key} requested={values[0]} status={result.Status} error={result.Error}");
                if (result.Status == DevicePowerPolicy.WriteStatus.Applied)
                    report.AddSuccess(component);
                else if (result.Status is DevicePowerPolicy.WriteStatus.NotExposed or DevicePowerPolicy.WriteStatus.NotSupported)
                    report.AddSkipped(component, "Live power management is not exposed by this device.", result.Error);
                else
                    report.AddError(component, "The restored live power setting could not be verified.",
                        $"status={result.Status} {result.Error}");
            }
            catch (Exception error)
            {
                report.AddError(component, "The restored live power setting could not be applied.", error.ToString());
            }
        }
        try { activatePowerScheme(); }
        catch (Exception error)
        {
            log($"BACKUP.RESTORE.POWER: power scheme activation failed: {error}");
            report.AddError("Power scheme", "The restored power scheme could not be activated.", error.ToString());
        }
        return report;
    }
}
