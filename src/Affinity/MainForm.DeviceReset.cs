using Microsoft.Win32;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void ApplyDeviceResetPlan(DeviceSettingsResetPlan plan, OperationReport? report)
    {
        DemandRealWrite("DEVICE SETTINGS");
        string regBase = plan.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        void RecordResetError(string operation, Exception ex)
        {
            report?.AddError($"{plan.DeviceName}: {operation}", ex.Message);
        }

        string prioPath = intBase + @"\Priority";
        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"RESET.REG.ERROR: {plan.InstanceId} operation=delete-priority-key path=HKLM\\{prioPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordResetError("remove legacy priority settings", ex);
        }

        string affPath = intBase + @"\Affinity Policy";
        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            if (key is not null)
            {
                foreach (string name in new[] { "DevicePriority", "DevicePolicy", "AssignmentSetOverride" })
                {
                    key.DeleteValue(name, throwOnMissingValue: false);
                }
            }
        }
        catch (Exception ex)
        {
            WriteLog($"RESET.REG.ERROR: {plan.InstanceId} operation=clear-affinity-values path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordResetError("clear affinity values", ex);
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(affPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"RESET.REG.ERROR: {plan.InstanceId} operation=delete-affinity-key path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordResetError("remove affinity settings key", ex);
        }

        if (plan.Kind == DeviceKind.NET_NDIS)
        {
            ClearNdisBaseCore(plan.InstanceId, report, plan.DeviceName);
            ClearNdisRssQueues(plan.InstanceId, report, plan.DeviceName);
            ClearNdisRssExtraValues(plan.InstanceId, report, plan.DeviceName);
        }

        ResetDevicePowerSaving(plan, report);

        WriteLog($"RESET: {plan.InstanceId} kind={plan.Kind} -> cleared priority/affinity (MSI left unchanged by design)");
    }

    private void ResetDevicePowerSaving(DeviceSettingsResetPlan plan, OperationReport? report)
    {
        DemandRealWrite("RESET DEVICE POWER");
        void RecordResetError(string operation, Exception ex)
        {
            report?.AddError($"{plan.DeviceName}: {operation}", ex.Message);
        }

        if (plan.Kind == DeviceKind.USB)
        {
            try
            {
                UsbSelectiveSuspendPolicy.ApplyResult powerResult =
                    UsbSelectiveSuspendPolicy.ApplyControllerAndHubs(plan.InstanceId, enabled: true);

                WriteLog(
                    $"RESET.POWER: {plan.InstanceId} -> Power Saving=Enabled " +
                    $"targets={powerResult.TargetCount} pnp={powerResult.PnpSucceeded}/{powerResult.TargetCount} " +
                    $"wmi={powerResult.WmiSucceeded}/{powerResult.WmiExposed} exposed");
                foreach (string note in powerResult.Notes)
                {
                    WriteLog($"RESET.POWER.NOTE: {note}");
                }
                if (!powerResult.Success)
                {
                    RecordResetError(
                        "reset USB power saving (Device Manager)",
                        new InvalidOperationException(string.Join(" | ", powerResult.Errors)));
                }
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.REG.ERROR: {plan.InstanceId} operation=reset-power-saving-usb error=\"{FlattenLogText(ex.ToString())}\"");
                RecordResetError("reset USB power saving", ex);
            }
        }

        if (plan.HasPowerEditor
            && (plan.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX)
            && !plan.Wifi)
        {
            try
            {
                string? classKey = GetClassKeyForDevice(plan.InstanceId);
                int pnpCaps = DevicePowerPolicy.ApplyNicPnPCapabilities(classKey ?? string.Empty, allowTurnOff: true);
                DevicePowerPolicy.WriteStatus msPowerStatus = DevicePowerPolicy.SetDevicePowerEnable(
                    plan.InstanceId,
                    allowTurnOff: true,
                    out string? msPowerError);
                WriteLog(
                    $"RESET.POWER: {plan.InstanceId} -> Power Saving=Enabled " +
                    $"PnPCapabilities=0x{pnpCaps:X} wmiWrite={msPowerStatus}");
                if (msPowerStatus is DevicePowerPolicy.WriteStatus.NotExposed or DevicePowerPolicy.WriteStatus.NotSupported)
                {
                    WriteLog($"RESET.POWER.NOTE: {plan.InstanceId}: MSPower_DeviceEnable status={msPowerStatus} {msPowerError}");
                }
                else if (DevicePowerPolicy.IsWriteFailure(msPowerStatus))
                {
                    WriteLog($"RESET.POWER.ERROR: {plan.InstanceId} wmiWrite={msPowerStatus} {msPowerError}");
                    RecordResetError(
                        "reset NIC power saving (Device Manager)",
                        new InvalidOperationException(msPowerError ?? "MSPower_DeviceEnable write failed"));
                }
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.REG.ERROR: {plan.InstanceId} operation=reset-power-saving-nic error=\"{FlattenLogText(ex.ToString())}\"");
                RecordResetError("reset NIC power saving", ex);
            }
        }
    }
}
