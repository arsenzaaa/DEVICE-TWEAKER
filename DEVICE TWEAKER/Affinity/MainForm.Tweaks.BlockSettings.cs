using Microsoft.Win32;
using System.Text;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private static void RecalcAffinityMask(DeviceBlock block)
    {
        ulong mask = 0;
        for (int i = 0; i < block.CpuBoxes.Count; i++)
        {
            if (block.CpuBoxes[i].Checked)
            {
                mask |= 1UL << i;
            }
        }

        block.AffinityMask = mask;
        if (block.Kind == DeviceKind.STOR)
        {
            block.AffinityLabel.Text = $"Affinity Mask: 0x{mask:X} (locked)";
        }
        else if (block.Kind == DeviceKind.NET_NDIS)
        {
            block.AffinityLabel.Text = $"Affinity (RSS mask): 0x{mask:X}";
        }
        else
        {
            block.AffinityLabel.Text = $"Affinity Mask: 0x{mask:X}";
        }

        // Keep Policy in sync with CPU selection so APPLY does what the checkboxes imply.
        if (block.Kind is DeviceKind.NET_NDIS or DeviceKind.STOR
            || !block.PolicyCombo.Enabled
            || block.PolicyCombo.Items.Count == 0)
        {
            return;
        }

        if (mask == 0)
        {
            if (block.PolicyCombo.Items.Contains("MachineDefault"))
            {
                block.PolicyCombo.SelectedItem = "MachineDefault";
            }

            return;
        }

        // MachineDefault ignores AssignmentSetOverride on APPLY; selecting CPUs means SpecCPU.
        if (string.Equals(block.PolicyCombo.SelectedItem?.ToString(), "MachineDefault", StringComparison.OrdinalIgnoreCase)
            && block.PolicyCombo.Items.Contains("SpecCPU"))
        {
            block.PolicyCombo.SelectedItem = "SpecCPU";
        }
    }

    private int ClampRssQueueCount(int value)
    {
        if (value < 1)
        {
            value = 1;
        }

        if (value > _maxLogical)
        {
            value = _maxLogical;
        }

        return value;
    }

    private int? GetFirstCheckedCore(DeviceBlock block)
    {
        for (int i = 0; i < block.CpuBoxes.Count; i++)
        {
            if (block.CpuBoxes[i].Checked)
            {
                return i;
            }
        }

        return null;
    }

    private NdisAffinityMode GetSelectedNdisAffinityMode(DeviceBlock block)
    {
        string mode = block.NdisModeCombo?.SelectedItem?.ToString() ?? "RSS";
        return mode.ToUpperInvariant() switch
        {
            "IRQ" => NdisAffinityMode.IrqPolicy,
            "BOTH" => NdisAffinityMode.Both,
            _ => NdisAffinityMode.Rss,
        };
    }

    private static string FormatNdisAffinityMode(NdisAffinityMode mode)
    {
        return mode switch
        {
            NdisAffinityMode.IrqPolicy => "IRQ",
            NdisAffinityMode.Both => "BOTH",
            _ => "RSS",
        };
    }

    private static string FormatNdisRuntimeValue(int? value)
    {
        return value.HasValue ? value.Value.ToString() : "-";
    }

    private static string FormatNdisRuntimeBool(bool? value)
    {
        return value.HasValue ? (value.Value ? "Enabled" : "Disabled") : "Unknown";
    }

    private static string FormatPowerSavingDisplay(string? stored)
    {
        if (string.Equals(stored, "off", StringComparison.OrdinalIgnoreCase))
        {
            return "Disabled";
        }

        if (string.Equals(stored, "on", StringComparison.OrdinalIgnoreCase))
        {
            return "Enabled";
        }

        return string.IsNullOrWhiteSpace(stored) ? "-" : stored.Trim();
    }

    private static string FormatNdisRssRuntimeState(NdisRssRuntimeState? state)
    {
        if (state is null)
        {
            return "active=Unknown";
        }

        return $"adapterFound={state.AdapterFound} rssFound={state.RssFound} active={FormatNdisRuntimeBool(state.Enabled)} adapter=\"{state.AdapterName}\" desc=\"{state.InterfaceDescription}\" base=G{FormatNdisRuntimeValue(state.BaseProcessorGroup)}:{FormatNdisRuntimeValue(state.BaseProcessorNumber)} max=G{FormatNdisRuntimeValue(state.MaxProcessorGroup)}:{FormatNdisRuntimeValue(state.MaxProcessorNumber)} maxProcessors={FormatNdisRuntimeValue(state.MaxProcessors)} queues={FormatNdisRuntimeValue(state.NumberOfReceiveQueues)} profile=\"{state.Profile}\" error=\"{state.Error}\"";
    }

    private static string BuildNdisRssConflictText(NdisRssRuntimeState? state, int? registryBase, int? registryQueues, int? registryMaxProcessors)
    {
        if (state is null)
        {
            return string.Empty;
        }

        bool registryConfigured = registryBase.HasValue || registryQueues.HasValue || registryMaxProcessors.HasValue;
        if (!state.RssFound)
        {
            return registryConfigured ? "registry RSS values exist but active RSS readback is unavailable" : string.Empty;
        }

        if (state.Enabled == false)
        {
            return registryConfigured ? "registry RSS values exist but active RSS is disabled" : string.Empty;
        }

        List<string> parts = [];
        if (registryBase.HasValue && state.BaseProcessorNumber.HasValue && registryBase.Value != state.BaseProcessorNumber.Value)
        {
            parts.Add($"base registry={registryBase.Value} active={state.BaseProcessorNumber.Value}");
        }

        if (registryQueues.HasValue && state.NumberOfReceiveQueues.HasValue && registryQueues.Value != state.NumberOfReceiveQueues.Value)
        {
            parts.Add($"queues registry={registryQueues.Value} active={state.NumberOfReceiveQueues.Value}");
        }

        if (registryMaxProcessors.HasValue && state.MaxProcessors.HasValue && registryMaxProcessors.Value != state.MaxProcessors.Value)
        {
            parts.Add($"maxProcessors registry={registryMaxProcessors.Value} active={state.MaxProcessors.Value}");
        }

        return parts.Count > 0 ? string.Join("; ", parts) : string.Empty;
    }

    private void LogNdisRssComparison(string prefix, string instanceId, NdisRssRuntimeState? state, int? registryBase, int? registryQueues, int? registryMaxProcessors)
    {
        string conflict = BuildNdisRssConflictText(state, registryBase, registryQueues, registryMaxProcessors);
        if (string.IsNullOrWhiteSpace(conflict))
        {
            WriteLog($"{prefix}.RSS.CHECK: {instanceId} registryBase={(registryBase?.ToString() ?? "-")} registryQueues={(registryQueues?.ToString() ?? "-")} registryMaxProcessors={(registryMaxProcessors?.ToString() ?? "-")} activeStatus=ok");
            return;
        }

        WriteLog($"{prefix}.RSS.CONFLICT: {instanceId} {conflict}");
    }

    private NdisAffinityMode ChooseSmartNdisAffinityMode(DeviceBlock block, int plannedQueues, out string reason)
    {
        NdisRssRuntimeState runtime = block.NdisRssRuntime ?? GetNdisRssRuntimeState(block.Device.InstanceId);
        block.NdisRssRuntime = runtime;

        bool runtimeRssKnown = runtime.RssFound;
        bool runtimeRssActive = runtime.RssFound && runtime.Enabled == true;
        bool registryRssConfigured = GetNdisBaseCore(block.Device.InstanceId).HasValue || GetNdisRssQueues(block.Device.InstanceId).HasValue;
        bool registryRssCapable = TestNdisRssBasePresent(block.Device.InstanceId);
        bool rssCapable = runtimeRssKnown || registryRssCapable || registryRssConfigured;
        bool msiEnabled = string.Equals(block.MsiCombo.SelectedItem?.ToString(), "Enabled", StringComparison.OrdinalIgnoreCase);
        bool multiQueue = plannedQueues > 1
            || (runtime.NumberOfReceiveQueues.HasValue && runtime.NumberOfReceiveQueues.Value > 1)
            || (runtime.MaxProcessors.HasValue && runtime.MaxProcessors.Value > 1);

        if (rssCapable && msiEnabled && (runtimeRssActive || registryRssConfigured))
        {
            reason = $"RSS capable + MSI enabled + {(runtimeRssActive ? "active RSS" : "registry RSS")} -> BOTH";
            return NdisAffinityMode.Both;
        }

        if (rssCapable && msiEnabled && multiQueue)
        {
            reason = "RSS capable + MSI enabled + multi-queue adapter -> BOTH";
            return NdisAffinityMode.Both;
        }

        if (rssCapable)
        {
            reason = msiEnabled
                ? "RSS capable + MSI enabled but no active/registry RSS yet -> RSS"
                : "RSS capable but MSI is not enabled yet -> RSS";
            return NdisAffinityMode.Rss;
        }

        reason = "RSS not detected -> IRQ";
        return NdisAffinityMode.IrqPolicy;
    }

    private void SetNdisModeCombo(DeviceBlock block, NdisAffinityMode mode)
    {
        if (block.NdisModeCombo is null)
        {
            return;
        }

        string text = FormatNdisAffinityMode(mode);
        if (block.NdisModeCombo.Items.Count == 0)
        {
            block.NdisModeCombo.Items.AddRange(["RSS", "IRQ", "BOTH"]);
        }

        block.NdisModeCombo.SelectedItem = text;
    }

    private static ulong ReadAffinityMaskValue(object? rawOverride)
    {
        return rawOverride switch
        {
            byte[] bytes when bytes.Length >= 8 => BitConverter.ToUInt64(bytes, 0),
            byte[] bytes when bytes.Length >= 4 => BitConverter.ToUInt32(bytes, 0),
            int intVal => (uint)intVal,
            uint uintVal => uintVal,
            long longVal => (ulong)longVal,
            ulong ulongVal => ulongVal,
            _ => 0,
        };
    }

    private static (int Policy, ulong Mask) ReadAffinityPolicyState(string affPath)
    {
        int policy = 0;
        ulong mask = 0;

        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath);
            if (affKey is null)
            {
                return (policy, mask);
            }

            if (affKey.GetValue("DevicePolicy") is int pv)
            {
                policy = pv;
            }

            mask = ReadAffinityMaskValue(affKey.GetValue("AssignmentSetOverride"));
        }
        catch
        {
        }

        return (policy, mask);
    }

    private void ApplyMaskToCpuBoxes(DeviceBlock block, ulong mask)
    {
        block.AffinityMask = mask;
        block.SuppressCpuEvents++;
        try
        {
            for (int i = 0; i < block.CpuBoxes.Count; i++)
            {
                ulong bit = 1UL << i;
                block.CpuBoxes[i].Checked = (mask & bit) != 0;
            }
        }
        finally
        {
            block.SuppressCpuEvents--;
        }

        RecalcAffinityMask(block);
    }

    private void WriteNdisIrqPolicy(DeviceBlock block, OperationReport? report = null)
    {
        string affPath = block.Device.RegBase + @"\Device Parameters\Interrupt Management\Affinity Policy";
        try
        {
            Registry.LocalMachine.CreateSubKey(affPath)?.Dispose();
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            if (affKey is null)
            {
                report?.AddError($"{block.Device.Name} — RSS IRQ policy", "registry key is unavailable");
                return;
            }

            ulong mask = block.AffinityMask;
            affKey.SetValue("DevicePolicy", 4, RegistryValueKind.DWord);
            byte[] bytes = IntPtr.Size >= 8 ? BitConverter.GetBytes(mask) : BitConverter.GetBytes((uint)mask);
            affKey.SetValue("AssignmentSetOverride", bytes, RegistryValueKind.Binary);
            WriteLog($"RSS.IRQ.SET: {block.Device.InstanceId} DevicePolicy=4 mask=0x{mask:X}");
        }
        catch (Exception ex)
        {
            WriteLog($"RSS.IRQ.SET: {block.Device.InstanceId} failed: {ex.Message}");
            report?.AddError($"{block.Device.Name} — RSS IRQ policy", ex.Message);
        }
    }

    private void ClearNdisIrqPolicy(DeviceBlock block, OperationReport? report = null)
    {
        string affPath = block.Device.RegBase + @"\Device Parameters\Interrupt Management\Affinity Policy";
        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            affKey?.DeleteValue("DevicePolicy", throwOnMissingValue: false);
            affKey?.DeleteValue("AssignmentSetOverride", throwOnMissingValue: false);
            WriteLog($"RSS.IRQ.CLEAR: {block.Device.InstanceId}");
        }
        catch (Exception ex)
        {
            WriteLog($"RSS.IRQ.CLEAR: {block.Device.InstanceId} failed: {ex.Message}");
            report?.AddError($"{block.Device.Name} — clear RSS IRQ policy", ex.Message);
        }
    }

    private void ApplyNdisSelection(DeviceBlock block, int baseCore, int queues)
    {
        int clampedQueues = ClampRssQueueCount(queues);
        int maxBase = Math.Max(0, _maxLogical - clampedQueues);
        if (baseCore < 0)
        {
            baseCore = 0;
        }
        else if (baseCore > maxBase)
        {
            baseCore = maxBase;
        }

        block.RssBaseCore = baseCore;

        if (block.RssQueueBox is not null && (int)block.RssQueueBox.Value != clampedQueues)
        {
            block.SuppressCpuEvents++;
            try
            {
                block.RssQueueBox.Value = clampedQueues;
            }
            finally
            {
                block.SuppressCpuEvents--;
            }
        }

        HashSet<int> selected = [];
        for (int i = 0; i < clampedQueues; i++)
        {
            selected.Add(baseCore + i);
        }

        block.SuppressCpuEvents++;
        try
        {
            foreach (CheckBox cb in block.CpuBoxes)
            {
                if (cb.Tag is not int core)
                {
                    continue;
                }
                bool isSelected = selected.Contains(core);
                cb.Checked = isSelected;
                cb.AutoCheck = !isSelected;
            }
        }
        finally
        {
            block.SuppressCpuEvents--;
        }

        RecalcAffinityMask(block);
    }

    private void HandleNdisCheckboxChanged(DeviceBlock block, CheckBox sender)
    {
        if (!sender.Checked)
        {
            return;
        }

        if (sender.Tag is not int baseCore)
        {
            return;
        }
        int queues = ClampRssQueueCount(block.RssQueueBox?.Value is decimal val ? (int)val : 1);
        ApplyNdisSelection(block, baseCore, queues);
    }

    private void LoadBlockSettings(DeviceBlock block)
    {
        block.SuppressCpuEvents++;
        try
        {
            foreach (CheckBox cb in block.CpuBoxes)
            {
                cb.Checked = false;
            }
        }
        finally
        {
            block.SuppressCpuEvents--;
        }

        block.AffinityMask = 0;

        bool isTestDevice = block.Device.IsTestDevice;
        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        string msiPath = intBase + @"\MessageSignaledInterruptProperties";

        int msiSupported = 0;
        int limit = 0;
        bool limitPresent = false;

        if (!isTestDevice)
        {
            try
            {
                using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath);
                if (msiKey is not null)
                {
                    msiSupported = msiKey.GetValue("MSISupported") as int? ?? 0;
                    if (msiKey.GetValue("MessageNumberLimit") is int limitVal)
                    {
                        limit = limitVal;
                        limitPresent = true;
                    }
                }
            }
            catch
            {
            }
        }
        else
        {
            // Test devices: honor explicit MSI status from Test Admin.
            msiSupported = block.Device.TestMsiStatus switch
            {
                "Enabled" => 1,
                "Disabled" => 0,
                _ => 1, // Auto / unknown -> Enabled for preview
            };
            limit = 0;
            limitPresent = true;
        }

        block.MsiCombo.SelectedItem = msiSupported == 1 ? "Enabled" : "Disabled";
        block.LimitBox.Text = limitPresent && limit > 0 ? limit.ToString() : "0";

        if (block.PowerSavingCheck is not null)
        {
            bool powerSavingEnabled = true;
            if (block.Kind == DeviceKind.USB)
            {
                if (!isTestDevice)
                {
                    _ = UsbSelectiveSuspendPolicy.TryReadEnabled(block.Device.InstanceId, out powerSavingEnabled);
                    block.Device.UsbSelectiveSuspend = powerSavingEnabled ? "on" : "off";
                }
                else
                {
                    powerSavingEnabled = !string.Equals(block.Device.UsbSelectiveSuspend, "off", StringComparison.OrdinalIgnoreCase);
                }
            }
            else if (!block.Device.Wifi)
            {
                if (!isTestDevice)
                {
                    bool? wmi = DevicePowerPolicy.TryReadDevicePowerEnable(block.Device.InstanceId);
                    if (wmi is bool live)
                    {
                        powerSavingEnabled = live;
                    }
                    else
                    {
                        int? pnpCaps = DevicePowerPolicy.TryReadNicPnPCapabilities(GetClassKeyForDevice(block.Device.InstanceId));
                        if (pnpCaps is int caps)
                        {
                            powerSavingEnabled = DevicePowerPolicy.IsNicTurnOffAllowed(caps);
                        }
                    }
                }
                else
                {
                    powerSavingEnabled = !string.Equals(block.Device.NicPowerSaving, "off", StringComparison.OrdinalIgnoreCase);
                }

                block.Device.NicPowerSaving = powerSavingEnabled ? "on" : "off";
            }

            block.PowerSavingCheck.Checked = powerSavingEnabled;
        }

        int? prioValue = null;
        string prioAffPath = intBase + @"\Affinity Policy";
        if (!isTestDevice)
        {
            try
            {
                using RegistryKey? prioKey = Registry.LocalMachine.OpenSubKey(prioAffPath);
                if (prioKey?.GetValue("DevicePriority") is int pv)
                {
                    prioValue = pv;
                }
            }
            catch
            {
            }
        }

        block.PrioCombo.SelectedItem = prioValue switch
        {
            1 => "Low",
            2 => "Normal",
            3 => "High",
            _ => "Undefined",
        };

        block.PolicyCombo.Items.Clear();
        if (block.Kind == DeviceKind.NET_NDIS)
        {
            block.PolicyLabel.Text = "Policy (RSS base)";
            FixRssPolicyLabelOverlap(block);
            block.PolicyCombo.Items.Add("RSS base core");
            block.PolicyCombo.SelectedIndex = 0;
            block.PolicyCombo.Enabled = false;
            block.PolicyLabel.Visible = false;
            block.PolicyCombo.Visible = false;

            bool skipTestWifiAffinity = isTestDevice && block.Device.Wifi;
            int? baseCore = skipTestWifiAffinity
                ? null
                : isTestDevice
                    ? 0
                    : GetNdisBaseCore(block.Device.InstanceId);
            int? registryQueues = isTestDevice ? 1 : GetNdisRssQueues(block.Device.InstanceId);
            int queues = registryQueues ?? 1;
            queues = ClampRssQueueCount(queues);
            block.NdisRssRuntime = isTestDevice
                ? new NdisRssRuntimeState(
                    AdapterFound: true,
                    RssFound: !skipTestWifiAffinity,
                    Enabled: !skipTestWifiAffinity,
                    BaseProcessorGroup: 0,
                    BaseProcessorNumber: baseCore ?? 0,
                    MaxProcessorGroup: 0,
                    MaxProcessorNumber: Math.Max(0, _maxLogical - 1),
                    MaxProcessors: Math.Max(1, _maxLogical),
                    NumberOfReceiveQueues: queues,
                    AdapterName: block.Device.Name,
                    InterfaceDescription: block.Device.Name,
                    Profile: "TEST",
                    Error: "test-device")
                : GetNdisRssRuntimeState(block.Device.InstanceId);
            (int irqPolicy, ulong irqMask) = isTestDevice
                ? (0, 0UL)
                : ReadAffinityPolicyState(prioAffPath);
            bool hasRss = baseCore.HasValue;
            bool hasIrqPolicy = irqPolicy == 4 && irqMask != 0;
            NdisAffinityMode ndisMode = skipTestWifiAffinity
                ? NdisAffinityMode.IrqPolicy
                : hasRss && hasIrqPolicy
                    ? NdisAffinityMode.Both
                    : hasIrqPolicy
                        ? NdisAffinityMode.IrqPolicy
                        : hasRss || block.NdisRssRuntime.RssFound
                            ? NdisAffinityMode.Rss
                            : NdisAffinityMode.IrqPolicy;
            SetNdisModeCombo(block, ndisMode);

            block.SuppressCpuEvents++;
            try
            {
                if (block.RssQueueBox is not null)
                {
                    block.RssQueueBox.Value = queues;
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            block.RssBaseCore = baseCore;
            if (skipTestWifiAffinity)
            {
                block.AffinityMask = 0;
            }
            else if (baseCore is >= 0 && baseCore < _maxLogical)
            {
                ApplyNdisSelection(block, baseCore.Value, queues);
            }
            else if (hasIrqPolicy)
            {
                int selectedCount = Math.Max(1, Enumerable.Range(0, block.CpuBoxes.Count).Count(i => (irqMask & (1UL << i)) != 0));
                if (block.RssQueueBox is not null)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        block.RssQueueBox.Value = ClampRssQueueCount(selectedCount);
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }
                }

                block.RssBaseCore = Enumerable.Range(0, block.CpuBoxes.Count).FirstOrDefault(i => (irqMask & (1UL << i)) != 0);
                ApplyMaskToCpuBoxes(block, irqMask);
            }

            string loadPrefix = isTestDevice ? "LOAD.TEST" : "LOAD";
            WriteLog($"{loadPrefix}.RSS.ACTIVE: {block.Device.InstanceId} {FormatNdisRssRuntimeState(block.NdisRssRuntime)}");
            LogNdisRssComparison(loadPrefix, block.Device.InstanceId, block.NdisRssRuntime, baseCore, registryQueues, null);
            WriteLog($"{loadPrefix}: NET_NDIS {block.Device.InstanceId} MSI={(msiSupported == 1 ? "Enabled" : "Disabled")} Limit={(limitPresent ? limit.ToString() : "Unlimited")} PrioVal={prioValue} Mode={FormatNdisAffinityMode(ndisMode)} BaseCore={(baseCore ?? -1)} Queues={queues} IrqPolicy={irqPolicy} IrqMask=0x{irqMask:X} Mask=0x{block.AffinityMask:X}");
        }
        else
        {
            block.PolicyLabel.Text = "Policy:";
            block.PolicyCombo.Items.AddRange(new object[] { "MachineDefault", "All", "AllClose", "Single", "SpecCPU", "SpreadMessages" });

            string affPath = intBase + @"\Affinity Policy";
            ulong mask = 0;
            int policyVal = 0;
            if (!isTestDevice)
            {
                try
                {
                    using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath);
                    if (affKey is not null)
                    {
                        if (affKey.GetValue("DevicePolicy") is int pv)
                        {
                            policyVal = pv;
                        }

                        mask = ReadAffinityMaskValue(affKey.GetValue("AssignmentSetOverride"));
                    }
                }
                catch
                {
                }
            }

            block.PolicyCombo.SelectedItem = policyVal switch
            {
                1 => "All",
                2 => "Single",
                3 => "AllClose",
                4 => "SpecCPU",
                5 => "SpreadMessages",
                _ => "MachineDefault",
            };

            block.AffinityMask = mask;
            block.SuppressCpuEvents++;
            try
            {
                for (int i = 0; i < block.CpuBoxes.Count; i++)
                {
                    ulong bit = 1UL << i;
                    block.CpuBoxes[i].Checked = (mask & bit) != 0;
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            string loadPrefix = isTestDevice ? "LOAD.TEST" : "LOAD";
            WriteLog($"{loadPrefix}: {block.Device.InstanceId} Kind={block.Kind} MSI={(msiSupported == 1 ? "Enabled" : "Disabled")} Limit={(limitPresent ? limit.ToString() : "Unlimited")} PrioVal={prioValue} PolicyVal={policyVal} Mask=0x{block.AffinityMask:X}");
        }

        RecalcAffinityMask(block);

        if (block.Kind == DeviceKind.USB)
        {
            if (IsUsbImodTarget(block.Device))
            {
                EnsureImodConfigLoaded();
                ImodConfig config = _imodConfigCache ?? new ImodConfig();
                string defaultText = FormatImodValue(config.GlobalInterval);
                block.ImodDefaultLabel.Text = $"default: {defaultText}";
                block.ImodDefaultLabel.Tag = defaultText;
                string currentText = block.ImodCurrentLabel.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(currentText)
                    || currentText.Equals("current: -", StringComparison.OrdinalIgnoreCase)
                    || currentText.Equals("current: reading...", StringComparison.OrdinalIgnoreCase))
                {
                    block.ImodCurrentLabel.Text = "current: unavailable";
                    block.ImodCurrentLabel.ForeColor = _statusInactive;
                }

                ImodConfigEntry? overrideEntry = FindImodOverride(block.Device.InstanceId, config);
                if (overrideEntry?.Enabled == false)
                {
                    block.ImodBox.Text = string.Empty;
                    block.ImodAutoCheck.Checked = false;
                }
                else
                {
                    bool hasCustomOverride = false;
                    if (overrideEntry?.RoleIntervals is { Count: > 0 } roleIntervals)
                    {
                        block.ImodBox.Text = FormatImodRoleIntervals(roleIntervals);
                        hasCustomOverride = roleIntervals.Values.Any(value => value != config.GlobalInterval);
                    }
                    else if (overrideEntry?.Intervals is { Count: > 0 } intervals)
                    {
                        block.ImodBox.Text = FormatImodVector(intervals);
                        hasCustomOverride = intervals.Any(value => value != config.GlobalInterval);
                    }
                    else
                    {
                        uint interval = GetEffectiveImodInterval(block.Device.InstanceId, config);
                        block.ImodBox.Text = FormatImodValue(interval);
                        hasCustomOverride = overrideEntry?.Interval.HasValue == true
                            && interval != config.GlobalInterval;
                    }

                    block.ImodAutoCheck.Checked = hasCustomOverride;
                }
            }
            else
            {
                block.ImodBox.Text = string.Empty;
                block.ImodAutoCheck.Checked = false;
                block.ImodDefaultLabel.Text = string.Empty;
                block.ImodDefaultLabel.Tag = null;
                block.ImodCurrentLabel.Text = string.Empty;
            }
        }
        else
        {
            block.ImodBox.Text = string.Empty;
            block.ImodAutoCheck.Checked = false;
            block.ImodDefaultLabel.Text = string.Empty;
            block.ImodDefaultLabel.Tag = null;
            block.ImodCurrentLabel.Text = string.Empty;
        }

        LoadRawMouseThrottleControls(block);
        RefreshNicItrBlock(block);
        UpdateImodSelectorsFromText(block);
        UpdateBlockInfoText(block);
    }

    private void UpdateBlockInfoText(
        DeviceBlock block,
        string? usbRolesOverride = null,
        string? usbPollingOverride = null,
        string? usbLivePollingOverride = null)
    {
        string shortPnp = GetShortPnpId(block.Device.InstanceId);
        string displayReg = GetDisplayRegPath(block.Device.InstanceId);
        string regBase = block.Device.RegBase;
        string usbRoles = usbRolesOverride ?? block.Device.UsbRoles;
        string usbPolling = usbPollingOverride ?? block.Device.UsbPollingRates;
        string usbLivePolling = usbLivePollingOverride ?? string.Empty;

        StringBuilder info = new();
        if (block.Device.IsTestDevice)
        {
            info.AppendLine("TEST DEVICE (no registry writes)");
        }
        info.AppendLine($"PNP ID: {shortPnp}");
        info.AppendLine($"Class: {block.Device.Class}");
        info.Append($"Registry: {displayReg}");

        if (block.Device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(usbRoles))
        {
            info.AppendLine();
            info.Append($"HID: {usbRoles}");
            if (!string.IsNullOrWhiteSpace(usbPolling))
            {
                info.AppendLine();
                info.Append($"Polling: {usbPolling}");
            }

            if (block.Device.UsbChipPath is UsbChipPathInfo chip)
            {
                info.AppendLine();
                info.Append(chip.DetailLine);
            }

            if (!string.IsNullOrWhiteSpace(block.Device.UsbSelectiveSuspend))
            {
                info.AppendLine();
                info.Append($"Power Saving: {FormatPowerSavingDisplay(block.Device.UsbSelectiveSuspend)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.USB)
        {
            if (block.Device.UsbChipPath is UsbChipPathInfo chip)
            {
                info.AppendLine();
                info.Append(chip.DetailLine);
            }

            if (!string.IsNullOrWhiteSpace(block.Device.UsbSelectiveSuspend))
            {
                info.AppendLine();
                info.Append($"Power Saving: {FormatPowerSavingDisplay(block.Device.UsbSelectiveSuspend)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.NET_NDIS)
        {
            info.AppendLine();
            string ndisMode = FormatNdisAffinityMode(GetSelectedNdisAffinityMode(block));
            info.Append($"Net type: NDIS ({ndisMode})");
            if (block.NdisRssRuntime is NdisRssRuntimeState runtime && runtime.RssFound)
            {
                info.AppendLine();
                info.Append($"RSS: {FormatNdisRuntimeBool(runtime.Enabled)} base {FormatNdisRuntimeValue(runtime.BaseProcessorNumber)} queues {FormatNdisRuntimeValue(runtime.NumberOfReceiveQueues)}");
            }
            if (TryGetNicItrProfile(block.Device.InstanceId) is NicItrProfile nicProfile)
            {
                info.AppendLine();
                info.Append($"NIC ITR: {nicProfile.FamilyName}");
            }

            if (!block.Device.Wifi && !string.IsNullOrWhiteSpace(block.Device.NicPowerSaving))
            {
                info.AppendLine();
                info.Append($"Power Saving: {FormatPowerSavingDisplay(block.Device.NicPowerSaving)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.NET_CX)
        {
            info.AppendLine();
            info.Append("Net type: NetAdapterCx");
            if (TryGetNicItrProfile(block.Device.InstanceId) is NicItrProfile nicProfile)
            {
                info.AppendLine();
                info.Append($"NIC ITR: {nicProfile.FamilyName}");
            }

            if (!block.Device.Wifi && !string.IsNullOrWhiteSpace(block.Device.NicPowerSaving))
            {
                info.AppendLine();
                info.Append($"Power Saving: {FormatPowerSavingDisplay(block.Device.NicPowerSaving)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.STOR)
        {
            info.AppendLine();
            info.Append("Type: Storage controller");
        }
        else if (block.Device.Kind == DeviceKind.AUDIO && !string.IsNullOrWhiteSpace(block.Device.AudioEndpoints))
        {
            info.AppendLine();
            info.Append($"Audio endpoints: {block.Device.AudioEndpoints}");
        }

        string infoText = info.ToString();
        if (!string.Equals(block.InfoLabel.Text, infoText, StringComparison.Ordinal))
        {
            block.InfoLabel.Text = infoText;
        }

        string fullRegPath = GetFullRegPath($"HKLM\\{regBase}");
        if (!string.Equals(block.InfoLabel.Tag as string, fullRegPath, StringComparison.Ordinal))
        {
            block.InfoLabel.Tag = fullRegPath;
        }
    }

    private void SaveBlockSettings(
        DeviceBlock block,
        bool msiOnlyForIntegratedGpu = false,
        OperationReport? report = null)
    {
        if (block.Device.IsTestDevice)
        {
            if (block.PowerSavingCheck is not null)
            {
                bool powerSavingEnabled = block.PowerSavingCheck.Checked;
                if (block.Kind == DeviceKind.USB)
                {
                    block.Device.UsbSelectiveSuspend = powerSavingEnabled ? "on" : "off";
                    WriteLog(
                        $"APPLY.PREVIEW: {block.Device.InstanceId} PowerSaving={block.Device.UsbSelectiveSuspend} " +
                        "(test USB — no registry/WMI/power plan)");
                }
                else if (!block.Device.Wifi)
                {
                    block.Device.NicPowerSaving = powerSavingEnabled ? "on" : "off";
                    WriteLog(
                        $"APPLY.PREVIEW: {block.Device.InstanceId} PowerSaving={block.Device.NicPowerSaving} " +
                        "(test NIC — no registry/WMI)");
                }
            }

            UpdateBlockInfoText(block);
            WriteLog($"APPLY.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason=TEST_DEVICE (preview only)");
            return;
        }

        RecalcAffinityMask(block);

        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        void RecordError(string operation, Exception ex)
        {
            report?.AddError($"{block.Device.Name} — {operation}", ex.Message);
        }

        try
        {
            Registry.LocalMachine.CreateSubKey(intBase)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=create-interrupt-key path=HKLM\\{intBase} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create interrupt settings key", ex);
        }

        string msiPath = intBase + @"\MessageSignaledInterruptProperties";
        try
        {
            Registry.LocalMachine.CreateSubKey(msiPath)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=create-msi-key path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create MSI settings key", ex);
        }

        string mode = block.MsiCombo.SelectedItem?.ToString() ?? "Disabled";
        int msiVal = mode == "Enabled" ? 1 : 0;
        try
        {
            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
            if (msiKey is null)
            {
                throw new InvalidOperationException("MSI registry key could not be opened for writing");
            }

            msiKey.SetValue("MSISupported", msiVal, RegistryValueKind.DWord);
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-msi value={msiVal} path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set MSI mode", ex);
        }

        if (msiOnlyForIntegratedGpu && block.Kind == DeviceKind.GPU && block.Device.IsIntegratedGpu)
        {
            WriteLog($"APPLY: {block.Device.InstanceId} MSI={mode} Kind={block.Kind} mode=autoIntegratedGpuMsiOnly");
            return;
        }

        string limitText = block.LimitBox.Text?.Trim() ?? string.Empty;
        bool isUnlimited = string.IsNullOrWhiteSpace(limitText) || limitText == "0" || Regex.IsMatch(limitText, "^(?i)unlimited$", RegexOptions.CultureInvariant);

        try
        {
            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
            if (msiKey is null)
            {
                throw new InvalidOperationException("MSI registry key could not be opened for writing");
            }

            if (isUnlimited)
            {
                msiKey.DeleteValue("MessageNumberLimit", throwOnMissingValue: false);
                block.LimitBox.Text = "0";
                limitText = "0";
            }
            else if (Regex.IsMatch(limitText, "^\\d+$", RegexOptions.CultureInvariant))
            {
                if (!int.TryParse(limitText, out int limitVal))
                {
                    throw new InvalidOperationException("MSI Limit is outside the supported 32-bit range");
                }

                if (limitVal < 0)
                {
                    limitVal = 0;
                }

                msiKey.SetValue("MessageNumberLimit", limitVal, RegistryValueKind.DWord);
                block.LimitBox.Text = limitVal.ToString();
                limitText = limitVal.ToString();
            }
            else
            {
                const string msiLimitMessage =
                    "MSI Limit must be a whole number. Leave empty or set 0 for unlimited. Value has been reset to 0 (unlimited).";
                if (report is null)
                {
                    ShowThemedInfo(msiLimitMessage);
                }
                else
                {
                    report.AddError($"{block.Device.Name} — MSI Limit", "invalid value; reset to 0 (unlimited)");
                }

                msiKey.DeleteValue("MessageNumberLimit", throwOnMissingValue: false);
                block.LimitBox.Text = "0";
                limitText = "0";
            }
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-msi-limit value=\"{limitText}\" path=HKLM\\{msiPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set MSI limit", ex);
        }

        string prioPath = intBase + @"\Priority";
        string prioAffPath = intBase + @"\Affinity Policy";
        try
        {
            Registry.LocalMachine.CreateSubKey(prioAffPath)?.Dispose();
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=create-affinity-key path=HKLM\\{prioAffPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create affinity settings key", ex);
        }

        string prioStr = block.PrioCombo.SelectedItem?.ToString() ?? "Undefined";
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
                prioKey.SetValue("DevicePriority", prioVal.Value, RegistryValueKind.DWord);
            }
            else
            {
                prioKey.DeleteValue("DevicePriority", throwOnMissingValue: false);
            }
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-priority value={prioStr} path=HKLM\\{prioAffPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set IRQ priority", ex);
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=delete-legacy-priority path=HKLM\\{prioPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("remove legacy priority settings", ex);
        }

        WriteLog($"APPLY: {block.Device.InstanceId} MSI={mode} Limit={limitText} Prio={prioStr} Mask=0x{block.AffinityMask:X} Kind={block.Kind}");

        if (block.PowerSavingCheck is not null && block.Kind == DeviceKind.USB)
        {
            bool suspendEnabled = block.PowerSavingCheck.Checked;
            string suspendMode = suspendEnabled ? "Enabled" : "Disabled";
            try
            {
                IReadOnlyList<string> hubs = UsbSelectiveSuspendPolicy.EnumerateRootHubs(block.Device.InstanceId);
                UsbSelectiveSuspendPolicy.ApplyControllerAndHubs(block.Device.InstanceId, suspendEnabled);
                bool msPowerOk = DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff: suspendEnabled);
                foreach (string hubId in hubs)
                {
                    msPowerOk |= DevicePowerPolicy.TrySetDevicePowerEnable(hubId, allowTurnOff: suspendEnabled);
                }

                block.Device.UsbSelectiveSuspend = suspendEnabled ? "on" : "off";
                WriteLog(
                    $"APPLY: {block.Device.InstanceId} PowerSaving={suspendMode} " +
                    $"controller+hubs={1 + hubs.Count} hubs=[{string.Join("; ", hubs)}] " +
                    $"(SelectiveSuspendEnabled + EnhancedPowerManagementEnabled + PnPCapabilities + MSPower) " +
                    $"wmiWrite={(msPowerOk ? "ok" : "miss")}");
                if (!msPowerOk)
                {
                    RecordError(
                        "set USB power saving (Device Manager)",
                        new InvalidOperationException("MSPower_DeviceEnable write missed; reboot may still be required"));
                }

                UpdateBlockInfoText(block);
            }
            catch (Exception ex)
            {
                WriteLog(
                    $"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-power-saving-usb " +
                    $"value={suspendMode} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("set USB power saving", ex);
            }
        }

        if (block.PowerSavingCheck is not null
            && (block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX)
            && !block.Device.Wifi)
        {
            bool allowTurnOff = block.PowerSavingCheck.Checked;
            string powerMode = allowTurnOff ? "Enabled" : "Disabled";
            try
            {
                string? classKey = GetClassKeyForDevice(block.Device.InstanceId);
                int pnpCaps = DevicePowerPolicy.ApplyNicPnPCapabilities(classKey ?? string.Empty, allowTurnOff);
                bool msPowerWrote = DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff);
                block.Device.NicPowerSaving = allowTurnOff ? "on" : "off";
                WriteLog(
                    $"APPLY: {block.Device.InstanceId} PowerSaving={powerMode} " +
                    $"PnPCapabilities=0x{pnpCaps:X} class=HKLM\\{classKey} " +
                    $"MSPower_DeviceEnable={(allowTurnOff ? "True" : "False")} wmiWrite={(msPowerWrote ? "ok" : "miss")}");
                if (!msPowerWrote)
                {
                    RecordError(
                        "set NIC power saving (Device Manager)",
                        new InvalidOperationException("MSPower_DeviceEnable write missed; reboot may still be required"));
                }

                UpdateBlockInfoText(block);
            }
            catch (Exception ex)
            {
                WriteLog(
                    $"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-power-saving-nic " +
                    $"value={powerMode} error=\"{FlattenLogText(ex.ToString())}\"");
                RecordError("set NIC power saving", ex);
            }
        }

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            int queues = ClampRssQueueCount(block.RssQueueBox?.Value is decimal val ? (int)val : 1);
            int baseCore = block.RssBaseCore ?? GetFirstCheckedCore(block) ?? 0;
            int maxBase = Math.Max(0, _maxLogical - queues);
            if (baseCore > maxBase)
            {
                baseCore = maxBase;
            }

            ApplyNdisSelection(block, baseCore, queues);

            NdisAffinityMode ndisMode = GetSelectedNdisAffinityMode(block);
            if (ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both)
            {
                SetNdisRssQueues(block.Device.InstanceId, queues, report, block.Device.Name);
                SetNdisBaseCore(block.Device.InstanceId, baseCore, report, block.Device.Name);
                SetNdisRssExtraValues(block.Device.InstanceId, baseCore, queues, report, block.Device.Name);
            }
            else
            {
                ClearNdisBaseCore(block.Device.InstanceId, report, block.Device.Name);
                ClearNdisRssQueues(block.Device.InstanceId, report, block.Device.Name);
                ClearNdisRssExtraValues(block.Device.InstanceId, report, block.Device.Name);
            }

            if (ndisMode is NdisAffinityMode.IrqPolicy or NdisAffinityMode.Both)
            {
                WriteNdisIrqPolicy(block, report);
            }
            else
            {
                ClearNdisIrqPolicy(block, report);
            }

            WriteLog($"APPLY: NET_NDIS {block.Device.InstanceId} mode={FormatNdisAffinityMode(ndisMode)} baseCore={baseCore} queues={queues} mask=0x{block.AffinityMask:X}");
            _ndisRssRuntimeCache.Remove(NormalizeInstanceId(block.Device.InstanceId));
            block.NdisRssRuntime = GetNdisRssRuntimeState(block.Device.InstanceId);
            WriteLog($"APPLY.RSS.ACTIVE: {block.Device.InstanceId} {FormatNdisRssRuntimeState(block.NdisRssRuntime)}");
            int? appliedRssBase = ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both ? baseCore : null;
            int? appliedRssQueues = ndisMode is NdisAffinityMode.Rss or NdisAffinityMode.Both ? queues : null;
            LogNdisRssComparison("APPLY", block.Device.InstanceId, block.NdisRssRuntime, appliedRssBase, appliedRssQueues, appliedRssQueues);
            return;
        }

        string affPath = intBase + @"\Affinity Policy";
        if (block.Kind == DeviceKind.STOR)
        {
            try
            {
                using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
                affKey?.DeleteValue("AssignmentSetOverride", throwOnMissingValue: false);
                affKey?.DeleteValue("DevicePolicy", throwOnMissingValue: false);
            }
            catch (Exception ex)
            {
                WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=clear-storage-affinity path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
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
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=create-affinity-key path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("create affinity policy key", ex);
        }

        string policyStr = block.PolicyCombo.SelectedItem?.ToString() ?? "MachineDefault";
        int policyVal = policyStr switch
        {
            "All" => 1,
            "Single" => 2,
            "AllClose" => 3,
            "SpecCPU" => 4,
            "SpreadMessages" => 5,
            _ => 0,
        };

        ulong mask = block.AffinityMask;
        if (policyVal == 0)
        {
            mask = 0;
        }

        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            if (affKey is null)
            {
                WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=open-affinity-key path=HKLM\\{affPath} error=key-unavailable");
                report?.AddError($"{block.Device.Name} — set affinity", "registry key is unavailable");
                return;
            }

            if (mask == 0 || policyVal == 0)
            {
                // Empty CPU selection / MachineDefault = clear DT affinity override (like RESET).
                affKey.SetValue("DevicePolicy", 0, RegistryValueKind.DWord);
                affKey.DeleteValue("AssignmentSetOverride", throwOnMissingValue: false);
                WriteLog($"APPLY: AFFINITY {block.Device.InstanceId} policy=MachineDefault value=0 mask cleared");
                return;
            }

            affKey.SetValue("DevicePolicy", policyVal, RegistryValueKind.DWord);
            byte[] bytes = IntPtr.Size >= 8 ? BitConverter.GetBytes(mask) : BitConverter.GetBytes((uint)mask);
            affKey.SetValue("AssignmentSetOverride", bytes, RegistryValueKind.Binary);
            WriteLog($"APPLY: AFFINITY {block.Device.InstanceId} policy={policyStr} value={policyVal} mask=0x{mask:X}");
        }
        catch (Exception ex)
        {
            WriteLog($"APPLY.REG.ERROR: {block.Device.InstanceId} operation=set-affinity policy={policyStr} mask=0x{mask:X} path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordError("set affinity", ex);
        }
    }

    private void ResetBlockSettings(DeviceBlock block, OperationReport? report = null)
    {
        if (block.Device.IsTestDevice)
        {
            block.SuppressCpuEvents++;
            try
            {
                foreach (CheckBox cb in block.CpuBoxes)
                {
                    cb.Checked = false;
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            block.AffinityMask = 0;
            block.AffinityLabel.Text = "Affinity Mask: 0x0";
            block.PrioCombo.SelectedItem = "Undefined";
            if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.RssBaseCore = null;
                if (block.RssQueueBox is not null)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        block.RssQueueBox.Value = 1;
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }
                }
            }
            else
            {
                block.PolicyCombo.SelectedItem = "MachineDefault";
            }

            block.IrqCount = null;
            block.IrqLabel.Text = "IRQ Count: reading...";
            if (block.PowerSavingCheck is not null)
            {
                block.PowerSavingCheck.Checked = true;
                if (block.Kind == DeviceKind.USB)
                {
                    block.Device.UsbSelectiveSuspend = "on";
                }
                else if (!block.Device.Wifi)
                {
                    block.Device.NicPowerSaving = "on";
                }
            }

            UpdateBlockInfoText(block);
            WriteLog($"RESET.TEST: {block.Device.InstanceId} kind={block.Kind} -> cleared preview priority/affinity/power");
            return;
        }

        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        void RecordResetError(string operation, Exception ex)
        {
            report?.AddError($"{block.Device.Name} — {operation}", ex.Message);
        }

        string prioPath = intBase + @"\Priority";
        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"RESET.REG.ERROR: {block.Device.InstanceId} operation=delete-priority-key path=HKLM\\{prioPath} error=\"{FlattenLogText(ex.ToString())}\"");
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
            WriteLog($"RESET.REG.ERROR: {block.Device.InstanceId} operation=clear-affinity-values path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordResetError("clear affinity values", ex);
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(affPath, throwOnMissingSubKey: false);
        }
        catch (Exception ex)
        {
            WriteLog($"RESET.REG.ERROR: {block.Device.InstanceId} operation=delete-affinity-key path=HKLM\\{affPath} error=\"{FlattenLogText(ex.ToString())}\"");
            RecordResetError("remove affinity settings key", ex);
        }

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            ClearNdisBaseCore(block.Device.InstanceId, report, block.Device.Name);
            ClearNdisRssQueues(block.Device.InstanceId, report, block.Device.Name);
            ClearNdisRssExtraValues(block.Device.InstanceId, report, block.Device.Name);
        }

        ResetDevicePowerSaving(block, report);

        WriteLog($"RESET: {block.Device.InstanceId} kind={block.Kind} -> cleared priority/affinity (MSI left unchanged by design)");
    }

    private void ResetDevicePowerSaving(DeviceBlock block, OperationReport? report)
    {
        void RecordResetError(string operation, Exception ex)
        {
            report?.AddError($"{block.Device.Name} — {operation}", ex.Message);
        }

        if (block.Kind == DeviceKind.USB)
        {
            try
            {
                UsbSelectiveSuspendPolicy.ApplyControllerAndHubs(block.Device.InstanceId, enabled: true);
                bool msPowerOk = DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff: true);
                foreach (string hubId in UsbSelectiveSuspendPolicy.EnumerateRootHubs(block.Device.InstanceId))
                {
                    msPowerOk |= DevicePowerPolicy.TrySetDevicePowerEnable(hubId, allowTurnOff: true);
                }

                block.Device.UsbSelectiveSuspend = "on";
                if (block.PowerSavingCheck is not null)
                {
                    block.PowerSavingCheck.Checked = true;
                }

                WriteLog(
                    $"RESET.POWER: {block.Device.InstanceId} -> Power Saving=Enabled " +
                    $"(USB controller+hubs + PnPCapabilities + MSPower) wmiWrite={(msPowerOk ? "ok" : "miss")}");
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.REG.ERROR: {block.Device.InstanceId} operation=reset-power-saving-usb error=\"{FlattenLogText(ex.ToString())}\"");
                RecordResetError("reset USB power saving", ex);
            }
        }

        if (block.PowerSavingCheck is not null
            && (block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX)
            && !block.Device.Wifi)
        {
            try
            {
                string? classKey = GetClassKeyForDevice(block.Device.InstanceId);
                int pnpCaps = DevicePowerPolicy.ApplyNicPnPCapabilities(classKey ?? string.Empty, allowTurnOff: true);
                bool msPowerWrote = DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff: true);
                block.PowerSavingCheck.Checked = true;
                block.Device.NicPowerSaving = "on";
                WriteLog(
                    $"RESET.POWER: {block.Device.InstanceId} -> Power Saving=Enabled " +
                    $"PnPCapabilities=0x{pnpCaps:X} wmiWrite={(msPowerWrote ? "ok" : "miss")}");
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.REG.ERROR: {block.Device.InstanceId} operation=reset-power-saving-nic error=\"{FlattenLogText(ex.ToString())}\"");
                RecordResetError("reset NIC power saving", ex);
            }
        }
    }

    private void ApplyUsbSelectiveSuspendPowerPlan(bool forceDisable, OperationReport? report)
    {
        // Fake USB blocks must not drive the real machine power-plan USB SS setting.
        // forceDisable (AUTO) also requires at least one real USB block — otherwise
        // test-only AUTO would still call SetPowerPlanEnabled(false) on the host.
        bool anyUsb = _blocks.Any(block =>
            block.Kind == DeviceKind.USB
            && !block.Device.IsTestDevice
            && block.PowerSavingCheck is not null);
        bool anyDisabled = _blocks.Any(block =>
            block.Kind == DeviceKind.USB
            && !block.Device.IsTestDevice
            && block.PowerSavingCheck is { Checked: false });
        if (!anyUsb)
        {
            WriteLog("USB.SUSPEND.PLAN: skipped (no real USB blocks)");
            return;
        }

        bool enable = !forceDisable && !anyDisabled;
        try
        {
            UsbSelectiveSuspendPolicy.SetPowerPlanEnabled(enable);
            WriteLog(
                $"USB.SUSPEND.PLAN: {(enable ? "enabled" : "disabled")} USB selective suspend for the active power scheme " +
                $"(subgroup={UsbSelectiveSuspendPolicy.UsbSettingsSubgroup:D} " +
                $"setting={UsbSelectiveSuspendPolicy.UsbSelectiveSuspendSetting:D} AC+DC={(enable ? 1 : 0)})");
        }
        catch (Exception ex)
        {
            WriteLog($"USB.SUSPEND.PLAN.ERROR: {FlattenLogText(ex.ToString())}");
            report?.AddError("USB Selective Suspend power plan", ex.Message);
        }
    }

    private void SyncLivePowerManagementAfterRestore()
    {
        foreach (DeviceBlock block in _blocks)
        {
            if (block.Device.IsTestDevice)
            {
                continue;
            }

            if (block.Kind == DeviceKind.USB)
            {
                _ = UsbSelectiveSuspendPolicy.TryReadEnabled(block.Device.InstanceId, out bool enabled);
                DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff: enabled);
                foreach (string hubId in UsbSelectiveSuspendPolicy.EnumerateRootHubs(block.Device.InstanceId))
                {
                    _ = UsbSelectiveSuspendPolicy.TryReadEnabled(hubId, out bool hubEnabled);
                    DevicePowerPolicy.TrySetDevicePowerEnable(hubId, allowTurnOff: hubEnabled);
                }
            }

            if (block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX && !block.Device.Wifi)
            {
                int? caps = DevicePowerPolicy.TryReadNicPnPCapabilities(GetClassKeyForDevice(block.Device.InstanceId));
                bool allowTurnOff = caps is not int value || DevicePowerPolicy.IsNicTurnOffAllowed(value);
                DevicePowerPolicy.TrySetDevicePowerEnable(block.Device.InstanceId, allowTurnOff);
            }
        }

        try
        {
            UsbSelectiveSuspendPolicy.ActivateCurrentPowerScheme();
            WriteLog("BACKUP.RESTORE: re-activated power scheme after USB/NIC power restore");
        }
        catch (Exception ex)
        {
            WriteLog($"BACKUP.RESTORE: power scheme activate failed: {FlattenLogText(ex.ToString())}");
        }
    }
}
