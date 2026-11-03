using Microsoft.Win32;
using System.Globalization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private void ApplyDeviceSettingsPlan(DeviceSettingsWritePlan plan, DeviceSettingsWriteResult result)
    {
        DemandRealWrite("DEVICE SETTINGS");
        OperationReport report = result.Report;
        string regBase = plan.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        void RecordError(string operation, Exception ex)
        {
            report?.AddError($"{plan.DeviceName}: {operation}", ex.Message);
        }

        try
        {
            Registry.LocalMachine.CreateSubKey(intBase)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=create-interrupt-key path=HKLM\\{intBase} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create interrupt settings key", ex);
        }

        string msiPath = intBase + @"\MessageSignaledInterruptProperties";
        try
        {
            Registry.LocalMachine.CreateSubKey(msiPath)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=create-msi-key path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create MSI settings key", ex);
        }

        string mode = plan.MsiMode;
        int msiVal = mode == "Enabled" ? 1 : 0;
        try
        {
            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
            if (msiKey is null)
            {
                throw new InvalidOperationException("MSI registry key could not be opened for writing");
            }

            DeviceSettingsWriter.SetMsi(new WindowsRegistryValues(msiKey), msiVal == 1);
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=set-msi value={msiVal} path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set MSI mode", ex);
        }

        if (plan.AutoMsiOnly)
        {
            string reason = plan.AutoReason;
            WriteLog($"APPLY: {plan.InstanceId} MSI={mode} Kind={plan.Kind} mode=autoMsiOnly reason={reason} preserved=limit,priority,policy,affinity");
            return;
        }

        string limitText = plan.LimitText;
        if (!TryValidateMsiLimitInput(limitText, plan.InterruptCapabilities, out bool isUnlocked, out int limitValue, out string validationError))
        {
            string msiLimitMessage = validationError + " The invalid value was not applied.";
            WriteLog($"APPLY.VALIDATION.ERROR: {plan.InstanceId} setting=MSI-Limit value=\"{FlattenLogText(limitText)}\" hardware=\"{FormatInterruptCapabilities(plan.InterruptCapabilities)}\" error=\"{FlattenLogText(validationError)}\"");
            report.AddError($"{plan.DeviceName}: MSI Limit", msiLimitMessage,
                $"Requested value={limitText}; hardware={FormatInterruptCapabilities(plan.InterruptCapabilities)}");
        }
        else
        {
            if (!isUnlocked && plan.InterruptCapabilities is null)
            {
                const string warning = "Windows did not publish this device's PCI interrupt capability; the numeric MSI limit could only be range-checked.";
                WriteLog($"APPLY.VALIDATION.WARN: {plan.InstanceId} setting=MSI-Limit value={limitValue} reason=capability-unknown");
                report?.AddWarning($"{plan.DeviceName}: MSI Limit", warning);
            }
            try
            {
                using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
                if (msiKey is null)
                {
                    throw new InvalidOperationException("MSI registry key could not be opened for writing");
                }

                if (isUnlocked)
                {
                    DeviceSettingsWriter.SetMsiLimit(new WindowsRegistryValues(msiKey), null);
                    result.NormalizedLimit = "Unlimited";
                    limitText = "Unlimited";
                }
                else
                {
                    DeviceSettingsWriter.SetMsiLimit(new WindowsRegistryValues(msiKey), limitValue);
                    result.NormalizedLimit = limitValue.ToString(CultureInfo.InvariantCulture);
                    limitText = limitValue.ToString(CultureInfo.InvariantCulture);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=set-msi-limit value=\"{limitText}\" path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("set MSI limit", ex);
            }
        }

        string prioPath = intBase + @"\Priority";
        string prioAffPath = intBase + @"\Affinity Policy";
        try
        {
            Registry.LocalMachine.CreateSubKey(prioAffPath)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=create-affinity-key path=HKLM\\{prioAffPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create affinity settings key", ex);
        }

        string prioStr = plan.Priority;
        int? prioVal = prioStr switch
        {
            "Low" => 1,
            "Normal" => 2,
            "High" => 3,
            _ => null,
        };

        try
        {
            using RegistryKey? prioKey = Registry.LocalMachine.OpenSubKey(prioAffPath, writable: true);
            if (prioKey is null)
            {
                throw new InvalidOperationException("Affinity registry key could not be opened for writing");
            }

            if (prioVal.HasValue)
            {
                DeviceSettingsWriter.SetPriority(new WindowsRegistryValues(prioKey), prioVal.Value);
            }
            else
            {
                DeviceSettingsWriter.SetPriority(new WindowsRegistryValues(prioKey), null);
            }
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=set-priority value={prioStr} path=HKLM\\{prioAffPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set IRQ priority", ex);
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=delete-legacy-priority path=HKLM\\{prioPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("remove legacy priority settings", ex);
        }

        WriteLog($"APPLY: {plan.InstanceId} MSI={mode} Limit={limitText} Prio={prioStr} Mask=0x{plan.AffinityMask:X} Kind={plan.Kind}");

        if (plan.PowerSaving.HasValue && plan.Kind == DeviceKind.USB)
        {
            bool suspendEnabled = plan.PowerSaving.Value;
            string suspendMode = suspendEnabled ? "Enabled" : "Disabled";
            try
            {
                UsbSelectiveSuspendPolicy.ApplyResult powerResult =
                    UsbSelectiveSuspendPolicy.ApplyControllerAndHubs(plan.InstanceId, suspendEnabled);

                result.UsbPower = suspendEnabled ? "on" : "off";
                WriteLog(
                    $"APPLY: {plan.InstanceId} PowerSaving={suspendMode} " +
                    $"controller+hubs={powerResult.TargetCount} " +
                    $"(SelectiveSuspendEnabled + EnhancedPowerManagementEnabled + PnPCapabilities + MSPower) " +
                    $"pnp={powerResult.PnpSucceeded}/{powerResult.TargetCount} " +
                    $"wmi={powerResult.WmiSucceeded}/{powerResult.WmiExposed} exposed");
                foreach (string note in powerResult.Notes)
                {
                    WriteLog($"APPLY.POWER.NOTE: {note}");
                }
                if (!powerResult.Success)
                {
                    RecordError(
                        "set USB power saving (Device Manager)",
                        new InvalidOperationException(string.Join(" | ", powerResult.Errors)));
                }

            }
            catch (Exception ex)
            {
                WriteLog(
                    $"APPLY.REG.ERROR: {plan.InstanceId} operation=set-power-saving-usb " +
                    $"value={suspendMode} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("set USB power saving", ex);
            }
        }

        if (plan.PowerSaving.HasValue
            && (plan.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX)
            && !plan.Wifi)
        {
            bool allowTurnOff = plan.PowerSaving.Value;
            string powerMode = allowTurnOff ? "Enabled" : "Disabled";
            try
            {
                string? classKey = GetClassKeyForDevice(plan.InstanceId);
                int pnpCaps = DevicePowerPolicy.ApplyNicPnPCapabilities(classKey ?? string.Empty, allowTurnOff);
                DevicePowerPolicy.WriteStatus msPowerStatus = DevicePowerPolicy.SetDevicePowerEnable(
                    plan.InstanceId,
                    allowTurnOff,
                    out string? msPowerError);
                result.NicPower = allowTurnOff ? "on" : "off";
                WriteLog(
                    $"APPLY: {plan.InstanceId} PowerSaving={powerMode} " +
                    $"PnPCapabilities=0x{pnpCaps:X} class=HKLM\\{classKey} " +
                    $"MSPower_DeviceEnable={(msPowerStatus == DevicePowerPolicy.WriteStatus.Applied ? allowTurnOff.ToString() : "unverified")} requested={allowTurnOff} wmiWrite={msPowerStatus}");
                if (msPowerStatus is DevicePowerPolicy.WriteStatus.NotExposed or DevicePowerPolicy.WriteStatus.NotSupported)
                {
                    WriteLog($"APPLY.POWER.NOTE: {plan.InstanceId}: MSPower_DeviceEnable status={msPowerStatus} {msPowerError}");
                }
                else if (DevicePowerPolicy.IsWriteFailure(msPowerStatus))
                {
                    WriteLog($"APPLY.POWER.ERROR: {plan.InstanceId} wmiWrite={msPowerStatus} {msPowerError}");
                    RecordError(
                        "set NIC power saving (Device Manager)",
                        new InvalidOperationException(msPowerError ?? "MSPower_DeviceEnable write failed"));
                }

            }
            catch (Exception ex)
            {
                WriteLog(
                    $"APPLY.REG.ERROR: {plan.InstanceId} operation=set-power-saving-nic " +
                    $"value={powerMode} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("set NIC power saving", ex);
            }
        }

        if (plan.Kind == DeviceKind.NET_NDIS)
        {
            NdisRssLayout.Layout layout = plan.RssLayout
                ?? throw new InvalidOperationException("Missing RSS layout in the write plan.");
            int baseCore = layout.BaseProcessor;
            int queues = layout.Processors.Count;
            NdisAffinityMode ndisMode = plan.NdisMode;
            if (ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both)
            {
                SetNdisRssQueues(plan.InstanceId, queues, report, plan.DeviceName);
                SetNdisBaseCore(plan.InstanceId, baseCore, report, plan.DeviceName);
                SetNdisRssExtraValues(plan.InstanceId, layout, report, plan.DeviceName);
            }
            else
            {
                ClearNdisBaseCore(plan.InstanceId, report, plan.DeviceName);
                ClearNdisRssQueues(plan.InstanceId, report, plan.DeviceName);
                ClearNdisRssExtraValues(plan.InstanceId, report, plan.DeviceName);
            }

            if (ndisMode is NdisAffinityMode.IrqPolicy or NdisAffinityMode.Both)
            {
                WriteNdisIrqPolicy(plan, report);
            }
            else
            {
                ClearNdisIrqPolicy(plan, report);
            }

            WriteLog($"APPLY: NET_NDIS {plan.InstanceId} mode={FormatNdisAffinityMode(ndisMode)} baseCore={baseCore} queues={queues} mask=0x{plan.AffinityMask:X}");
            result.RssRuntime = ReadNdisRssRuntimeState(plan.InstanceId);
            WriteLog($"APPLY.RSS.ACTIVE: {plan.InstanceId} {FormatNdisRssRuntimeState(result.RssRuntime)}");
            int? appliedRssBase = ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both ? baseCore : null;
            int? appliedRssQueues = ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both ? queues : null;
            LogNdisRssComparison("APPLY", plan.InstanceId, result.RssRuntime, appliedRssBase, appliedRssQueues, appliedRssQueues);
            return;
        }

        string affPath = intBase + @"\Affinity Policy";
        if (plan.Kind == DeviceKind.STOR)
        {
            try
            {
                using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
                if (affKey is not null) DeviceSettingsWriter.ClearStorageAffinity(new WindowsRegistryValues(affKey));
                WriteLog(
                    $"APPLY: STORAGE {plan.InstanceId} MSI={mode} Limit={limitText} Prio={prioStr} " +
                    "affinity=WindowsDefault removed=AssignmentSetOverride,DevicePolicy");
            }
            catch (Exception ex)
            {
                WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=clear-storage-affinity path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("clear storage affinity", ex);
            }

            return;
        }

        try
        {
            Registry.LocalMachine.CreateSubKey(affPath)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=create-affinity-key path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create affinity policy key", ex);
        }

        string policyStr = plan.Policy;
        int policyVal = MapPolicyText(policyStr) ?? 0;

        ulong mask = plan.AffinityMask;
        if (policyVal == 0)
        {
            mask = 0;
        }

        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            if (affKey is null)
            {
                WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=open-affinity-key path=HKLM\\{affPath} error=key-unavailable");
                report?.AddError($"{plan.DeviceName}: set affinity", "registry key is unavailable");
                return;
            }

            if (mask == 0 || policyVal == 0)
            {
                // Empty CPU selection / MachineDefault = clear DT affinity override (like RESET).
                DeviceSettingsWriter.SetAffinity(new WindowsRegistryValues(affKey), policyVal, mask, IntPtr.Size);
                WriteLog($"APPLY: AFFINITY {plan.InstanceId} policy=MachineDefault value=0 mask cleared");
                return;
            }

            DeviceSettingsWriter.SetAffinity(new WindowsRegistryValues(affKey), policyVal, mask, IntPtr.Size);
            WriteLog($"APPLY: AFFINITY {plan.InstanceId} policy={policyStr} value={policyVal} mask=0x{mask:X}");
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {plan.InstanceId} operation=set-affinity policy={policyStr} mask=0x{mask:X} path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set affinity", ex);
        }
    }

}
