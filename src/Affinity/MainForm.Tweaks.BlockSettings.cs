using System.Globalization;
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
            block.AffinityLabel.Text = "Affinity Mask: Windows Default";
        }
        else if (block.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(block.Device.InstanceId, block.Device.Name) || IsDisplayAudioEndpointsText(block.Device.AudioEndpoints)) && mask == 0)
        {
            block.AffinityLabel.Text = "Affinity Mask: 0x0 (Windows Default)";
        }
        else if (block.Kind == DeviceKind.NET_NDIS)
        {
            string ndisMode = block.NdisModeCombo?.SelectedItem?.ToString() ?? "RSS";
            block.AffinityLabel.Text = string.Equals(ndisMode, "IRQ", StringComparison.OrdinalIgnoreCase)
                ? $"Affinity Mask: 0x{mask:X}"
                : $"Affinity (RSS mask): 0x{mask:X}";
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

        int policy = MapPolicyText(block.PolicyCombo.SelectedItem?.ToString() ?? string.Empty) ?? 0;
        string presented = FormatPolicyValue(NormalizePolicyForSelection(policy, mask));
        if (block.PolicyCombo.Items.Contains(presented))
        {
            block.PolicyCombo.SelectedItem = presented;
        }
    }

    internal static int NormalizePolicyForSelection(int policy, ulong mask)
        => mask == 0 ? 0 : policy == 0 ? 4 : policy;

    private int ClampRssQueueCount(int value)
    {
        if (value < 1)
        {
            value = 1;
        }

        int maxQueues = NdisRssLayout.GetMaxQueues(_cpuInfo?.Topology.LPs, _maxLogical);
        if (value > maxQueues)
        {
            value = maxQueues;
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
        return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "-";
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

        return $"adapterFound={state.AdapterFound} rssFound={state.RssFound} active={FormatNdisRuntimeBool(state.Enabled)} readStatus={NdisRssDiagnostics.ReadStatus(state)} "
            + $"adapter=\"{SanitizeLogValue(state.AdapterName)}\" desc=\"{SanitizeLogValue(state.InterfaceDescription)}\" "
            + $"base=G{FormatNdisRuntimeValue(state.BaseProcessorGroup)}:{FormatNdisRuntimeValue(state.BaseProcessorNumber)} "
            + $"max=G{FormatNdisRuntimeValue(state.MaxProcessorGroup)}:{FormatNdisRuntimeValue(state.MaxProcessorNumber)} "
            + $"maxProcessors={FormatNdisRuntimeValue(state.MaxProcessors)} queues={FormatNdisRuntimeValue(state.NumberOfReceiveQueues)} "
            + $"profile=\"{SanitizeLogValue(state.Profile)}\" error=\"{SanitizeLogValue(state.Error)}\"";
    }

    private void LogNdisRssComparison(string prefix, string instanceId, NdisRssRuntimeState? state, int? registryBase, int? registryQueues, int? registryMaxProcessors, bool wifiPreserved = false)
    {
        NdisRssDiagnostics.Comparison result = NdisRssDiagnostics.Compare(state, registryBase, registryQueues, registryMaxProcessors, wifiPreserved);
        string eventName = result.Status == "different" ? "CONFLICT" : result.Status == "failed" ? "ERROR" : "CHECK";
        WriteLog($"{prefix}.RSS.{eventName}: {instanceId} registryBase={(registryBase?.ToString(CultureInfo.InvariantCulture) ?? "-")} registryQueues={(registryQueues?.ToString(CultureInfo.InvariantCulture) ?? "-")} registryMaxProcessors={(registryMaxProcessors?.ToString(CultureInfo.InvariantCulture) ?? "-")} activeStatus={result.Status} detail=\"{SanitizeLogValue(result.Detail)}\"");
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

    internal static ulong ReadAffinityMaskValue(object? rawOverride)
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
        catch (Exception ex)
        {
            AppDiagnostics.Warn("AFFINITY.READ", ex);
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

    private void WriteNdisIrqPolicy(DeviceSettingsWritePlan plan, OperationReport? report = null)
    {
        DemandRealWrite("WriteNdisIrqPolicy");
        string affPath = plan.RegBase + @"\Device Parameters\Interrupt Management\Affinity Policy";
        try
        {
            Registry.LocalMachine.CreateSubKey(affPath)?.Dispose();
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            if (affKey is null)
            {
                report?.AddError($"{plan.DeviceName}: RSS IRQ policy", "registry key is unavailable");
                return;
            }

            ulong mask = plan.AffinityMask;
            affKey.SetValue("DevicePolicy", 4, RegistryValueKind.DWord);
            byte[] bytes = IntPtr.Size >= 8 ? BitConverter.GetBytes(mask) : BitConverter.GetBytes((uint)mask);
            affKey.SetValue("AssignmentSetOverride", bytes, RegistryValueKind.Binary);
            WriteLog($"RSS.IRQ.SET: {plan.InstanceId} DevicePolicy=4 mask=0x{mask:X}");
        }
        catch (Exception ex)
        {
            WriteLog($"RSS.IRQ.SET: {plan.InstanceId} failed: {ex.Message}");
            report?.AddError($"{plan.DeviceName}: RSS IRQ policy", ex.Message);
        }
    }

    private void ClearNdisIrqPolicy(DeviceSettingsWritePlan plan, OperationReport? report = null)
    {
        DemandRealWrite("ClearNdisIrqPolicy");
        string affPath = plan.RegBase + @"\Device Parameters\Interrupt Management\Affinity Policy";
        try
        {
            using RegistryKey? affKey = Registry.LocalMachine.OpenSubKey(affPath, writable: true);
            affKey?.DeleteValue("DevicePolicy", throwOnMissingValue: false);
            affKey?.DeleteValue("AssignmentSetOverride", throwOnMissingValue: false);
            WriteLog($"RSS.IRQ.CLEAR: {plan.InstanceId}");
        }
        catch (Exception ex)
        {
            WriteLog($"RSS.IRQ.CLEAR: {plan.InstanceId} failed: {ex.Message}");
            report?.AddError($"{plan.DeviceName}: clear RSS IRQ policy", ex.Message);
        }
    }

    private NdisRssLayout.Layout GetNdisRssLayout(int baseCore, int queues)
        => NdisRssLayout.Select(_cpuInfo?.Topology.LPs, _maxLogical, Math.Max(0, baseCore), ClampRssQueueCount(queues));

    private void UpdateRssQueueOptions(DeviceBlock block)
    {
        // Wi-Fi is never written, so its card shows the stored queue count as is
        // instead of snapping it to a value APPLY would use.
        if (block.Kind != DeviceKind.NET_NDIS || block.Device.Wifi)
        {
            return;
        }

        List<int> driverValues = block.Device.IsTestDevice ? [] : GetNdisRssQueueDriverValues(block.Device.InstanceId);
        List<int> options = NdisRssLayout.GetAllowedQueueCounts(
            driverValues,
            block.InterruptCapabilities,
            NdisRssLayout.GetMaxQueues(_cpuInfo?.Topology.LPs, _maxLogical));
        block.RssQueueOptions = options;
        WriteLog($"RSS.QUEUES.OPTIONS: {block.Device.InstanceId} driver=[{string.Join(',', driverValues)}] hw={FormatInterruptCapabilities(block.InterruptCapabilities)} allowed=[{string.Join(',', options)}]");

        if (block.RssQueueBox is null)
        {
            return;
        }

        block.SuppressCpuEvents++;
        try
        {
            block.RssQueueBox.Maximum = options[^1];
        }
        finally
        {
            block.SuppressCpuEvents--;
        }

        int current = (int)block.RssQueueBox.Value;
        if (!options.Contains(current) && block.RssBaseCore is int baseCore)
        {
            ApplyNdisSelection(block, baseCore, current);
        }
    }

    private int FloorRssQueueCount(DeviceBlock block, int queues)
    {
        int clamped = ClampRssQueueCount(queues);
        return block.RssQueueOptions is { Count: > 0 } options ? NdisRssLayout.Floor(options, clamped) : clamped;
    }

    private NdisRssLayout.Layout ApplyNdisSelection(DeviceBlock block, int baseCore, int queues)
    {
        int clampedQueues = FloorRssQueueCount(block, queues);
        NdisRssLayout.Layout layout = GetNdisRssLayout(baseCore, clampedQueues);
        clampedQueues = layout.Processors.Count;
        block.RssBaseCore = layout.BaseProcessor;
        block.RssQueueLast = clampedQueues;

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

        HashSet<int> selected = [.. layout.Processors];

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
        return layout;
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
            catch (Exception ex)
            {
                AppDiagnostics.Warn("MSI.READ", ex);
            }
        }
        else
        {
            TestDeviceState testState = EnsureTestDeviceState(block.Device);
            msiSupported = testState.MsiEnabled ? 1 : 0;
            limit = testState.MsiLimit ?? 0;
            limitPresent = testState.MsiLimit.HasValue;
        }

        block.MsiCombo.SelectedItem = msiSupported == 1 ? "Enabled" : "Disabled";
        block.LimitBox.Text = limitPresent && limit > 0 ? limit.ToString(CultureInfo.InvariantCulture) : "Unlimited";

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
                    powerSavingEnabled = EnsureTestDeviceState(block.Device).PowerSavingEnabled
                        ?? !string.Equals(block.Device.UsbSelectiveSuspend, "off", StringComparison.OrdinalIgnoreCase);
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
                    powerSavingEnabled = EnsureTestDeviceState(block.Device).PowerSavingEnabled
                        ?? !string.Equals(block.Device.NicPowerSaving, "off", StringComparison.OrdinalIgnoreCase);
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
            catch (Exception ex)
            {
                AppDiagnostics.Warn("PRIORITY.READ", ex);
            }
        }
        else
        {
            prioValue = EnsureTestDeviceState(block.Device).Priority;
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
            TestDeviceState? testState = isTestDevice ? EnsureTestDeviceState(block.Device) : null;
            int? baseCore = skipTestWifiAffinity
                ? null
                : isTestDevice
                    ? testState!.RssBaseCore
                    : GetNdisBaseCore(block.Device.InstanceId);
            int? registryQueues = isTestDevice ? testState!.RssQueues : GetNdisRssQueues(block.Device.InstanceId);
            UpdateRssQueueOptions(block);
            int queues = FloorRssQueueCount(block, registryQueues ?? 1);
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
                    Error: string.Empty)
                : block.Device.Wifi ? EmptyNdisRssRuntimeState("wifi-preserved") : GetNdisRssRuntimeState(block.Device.InstanceId);
            (int irqPolicy, ulong irqMask) = isTestDevice
                ? (testState!.Policy, testState.AffinityMask)
                : ReadAffinityPolicyState(prioAffPath);
            bool hasRss = baseCore.HasValue;
            bool hasIrqPolicy = irqPolicy == 4 && irqMask != 0;
            NdisAffinityMode ndisMode = isTestDevice && !skipTestWifiAffinity
                ? testState!.NdisMode.ToUpperInvariant() switch
                {
                    "IRQ" => NdisAffinityMode.IrqPolicy,
                    "BOTH" => NdisAffinityMode.Both,
                    _ => NdisAffinityMode.Rss,
                }
                : skipTestWifiAffinity
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
                        block.RssQueueBox.Value = FloorRssQueueCount(block, selectedCount);
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
            LogNdisRssComparison(loadPrefix, block.Device.InstanceId, block.NdisRssRuntime, baseCore, registryQueues, null, block.Device.Wifi);
            WriteLog($"{loadPrefix}: NET_NDIS {block.Device.InstanceId} MSI={(msiSupported == 1 ? "Enabled" : "Disabled")} Limit={(limitPresent ? limit.ToString(CultureInfo.InvariantCulture) : "Unlimited")} PrioVal={prioValue} Mode={FormatNdisAffinityMode(ndisMode)} BaseCore={(baseCore ?? -1)} Queues={queues} IrqPolicy={irqPolicy} IrqMask=0x{irqMask:X} Mask=0x{block.AffinityMask:X}");
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
                catch (Exception ex)
                {
                    AppDiagnostics.Warn("AFFINITY.READ", ex);
                }
            }
            else
            {
                TestDeviceState testState = EnsureTestDeviceState(block.Device);
                policyVal = testState.Policy;
                mask = testState.AffinityMask;
            }

            block.PolicyCombo.SelectedItem = FormatPolicyValue(policyVal);

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
            WriteLog($"{loadPrefix}: {block.Device.InstanceId} Kind={block.Kind} MSI={(msiSupported == 1 ? "Enabled" : "Disabled")} Limit={(limitPresent ? limit.ToString(CultureInfo.InvariantCulture) : "Unlimited")} PrioVal={prioValue} PolicyVal={policyVal} Mask=0x{block.AffinityMask:X}");
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
                if (config.StartupReadError is not null)
                {
                    block.ImodDefaultLabel.Text = "profile: unreadable";
                    block.ImodDefaultLabel.Tag = config.StartupReadError;
                    block.ImodDefaultLabel.ForeColor = _statusDanger;
                    SetImodStatusTooltip(block.ImodDefaultLabel);
                }
                string currentText = GetSourceControlText(block.ImodCurrentLabel);
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

        if (isTestDevice && IsUsbImodTarget(block.Device))
            block.ImodBox.Text = EnsureTestDeviceState(block.Device).ImodValue;

        LoadRawMouseThrottleControls(block);
        UpdateImodSelectorsFromText(block);
        UpdateBlockInfoText(block);
    }

    private void AppendNicItrInfo(StringBuilder info, DeviceBlock block)
    {
        if (TryGetNicItrProfile(block.Device.InstanceId) is not NicItrProfile nicProfile)
        {
            return;
        }

        info.AppendLine();
        info.Append(CultureInfo.InvariantCulture, $"NIC ITR: {nicProfile.FamilyName}");
        if (!block.Device.IsTestDevice
            && NicDriverModeration.Describe(ReadNicDriverModeration(block.Device.InstanceId)) is string moderation)
        {
            info.AppendLine();
            info.Append(CultureInfo.InvariantCulture, $"Interrupt Moderation: {moderation}");
        }
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
        if (block.Device.Kind == DeviceKind.USB)
        {
            string controllerKey = NormalizeInstanceId(block.Device.InstanceId);
            if (usbRolesOverride is null
                && _rawPollingByController.TryGetValue(controllerKey, out Dictionary<string, string>? measuredRoles)
                && measuredRoles.Count > 0)
            {
                usbRoles = ApplyUsbRolePollingOverrides(usbRoles, measuredRoles);
                if (usbPollingOverride is null)
                {
                    usbPolling = FormatUsbPollingRoleSummary(usbRoles.Split(',', StringSplitOptions.RemoveEmptyEntries));
                }
            }

            if (usbLivePollingOverride is null
                && _rawLivePollingByController.TryGetValue(controllerKey, out Dictionary<string, string>? liveRoles)
                && liveRoles.Count > 0)
            {
                usbLivePolling = FormatLivePollingSummary(liveRoles);
            }
        }

        StringBuilder info = new();
        if (block.Device.IsTestDevice && Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_HIDE_SANDBOX_HEADER") != "1")
        {
            info.AppendLine("TEST DEVICE (no registry writes)");
        }
        info.AppendLine(CultureInfo.InvariantCulture, $"PNP ID: {shortPnp}");
        info.AppendLine(CultureInfo.InvariantCulture, $"Class: {block.Device.Class}");
        info.Append(CultureInfo.InvariantCulture, $"Registry: {displayReg}");

        if (block.Device.Wifi)
        {
            info.AppendLine();
            info.Append("Protection: Wi-Fi settings preserved (no changes)");
        }

        if (block.Device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(usbRoles))
        {
            info.AppendLine();
            info.Append(CultureInfo.InvariantCulture, $"HID: {usbRoles}");
            if (!string.IsNullOrWhiteSpace(usbPolling))
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"Polling: {usbPolling}");
            }
            if (!string.IsNullOrWhiteSpace(usbLivePolling))
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"Polling live: {usbLivePolling}");
            }
            if (Regex.IsMatch($"{usbRoles} {usbPolling} {usbLivePolling}", @"(?i)(?:[4-9](?:\.\d+)?\s*k(?:hz)?|(?:4000|8000)\s*hz)"))
            {
                info.AppendLine();
                info.Append("AUTO profile: Mouse IMOD = 0 (interrupt moderation off). Manual IMOD is set separately.");
            }

            if (block.Device.UsbChipPath is UsbChipPathInfo chip)
            {
                info.AppendLine();
                info.Append(chip.DetailLine);
            }

            if (!string.IsNullOrWhiteSpace(block.Device.UsbSelectiveSuspend))
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"Power Saving: {FormatPowerSavingDisplay(block.Device.UsbSelectiveSuspend)}");
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
                info.Append(CultureInfo.InvariantCulture, $"Power Saving: {FormatPowerSavingDisplay(block.Device.UsbSelectiveSuspend)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.NET_NDIS)
        {
            info.AppendLine();
            string ndisMode = FormatNdisAffinityMode(GetSelectedNdisAffinityMode(block));
            info.Append(CultureInfo.InvariantCulture, $"Net type: NDIS ({ndisMode})");
            if (block.NdisRssRuntime is NdisRssRuntimeState runtime && runtime.RssFound)
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"RSS: {FormatNdisRuntimeBool(runtime.Enabled)} base {FormatNdisRuntimeValue(runtime.BaseProcessorNumber)} queues {FormatNdisRuntimeValue(runtime.NumberOfReceiveQueues)}");
            }
            AppendNicItrInfo(info, block);

            if (!block.Device.Wifi && !string.IsNullOrWhiteSpace(block.Device.NicPowerSaving))
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"Power Saving: {FormatPowerSavingDisplay(block.Device.NicPowerSaving)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.NET_CX)
        {
            info.AppendLine();
            info.Append("Net type: NetAdapterCx");
            AppendNicItrInfo(info, block);

            if (!block.Device.Wifi && !string.IsNullOrWhiteSpace(block.Device.NicPowerSaving))
            {
                info.AppendLine();
                info.Append(CultureInfo.InvariantCulture, $"Power Saving: {FormatPowerSavingDisplay(block.Device.NicPowerSaving)}");
            }
        }
        else if (block.Device.Kind == DeviceKind.STOR)
        {
            info.AppendLine();
            bool isNvme = Regex.IsMatch(block.Device.Name ?? string.Empty, "(?i)NVM\\s*Express|NVMe")
                || string.Equals(block.Device.StorageTag, "NVMe", StringComparison.OrdinalIgnoreCase);
            bool isSata = Regex.IsMatch(block.Device.Name ?? string.Empty, "(?i)\\bSATA\\b|\\bAHCI\\b");
            if (isNvme)
            {
                info.Append("Type: NVMe storage controller");
            }
            else if (isSata)
            {
                info.Append("Type: SATA/AHCI controller");
            }
            else
            {
                info.Append("Type: Storage controller");
            }
        }
        else if (block.Device.Kind == DeviceKind.AUDIO && !string.IsNullOrWhiteSpace(block.Device.AudioEndpoints))
        {
            info.AppendLine();
            info.Append(CultureInfo.InvariantCulture, $"Audio endpoints: {block.Device.AudioEndpoints}");
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

    // Keep the real-write guard synchronous and before any UI/state access.
    private Task SaveBlockSettings(DeviceBlock block, bool autoMsiOnly = false, OperationReport? report = null)
    {
        if (block.Device.IsTestDevice)
        {
            ApplyTestBlockSettings(block, autoMsiOnly, report);
            UpdateBlockInfoText(block);
            WriteLog($"APPLY.VIRTUAL: {block.Device.InstanceId} Kind={block.Kind} reason=TEST_DEVICE");
            return Task.CompletedTask;
        }
        DemandRealWrite("DEVICE SETTINGS");
        RecalcAffinityMask(block);
        NdisRssLayout.Layout? layout = block.Kind == DeviceKind.NET_NDIS && !autoMsiOnly
            ? ApplyNdisSelection(block, block.RssBaseCore ?? GetFirstCheckedCore(block) ?? 0,
                ClampRssQueueCount((int)(block.RssQueueBox?.Value ?? 1))) : null;
        if (layout is not null) layout = layout with { Processors = Array.AsReadOnly(layout.Processors.ToArray()) };
        DeviceSettingsWritePlan plan = new()
        {
            InstanceId = block.Device.InstanceId, DeviceName = block.Device.Name, RegBase = block.Device.RegBase,
            Kind = block.Kind, MsiMode = block.MsiCombo.SelectedItem?.ToString() ?? "Disabled",
            LimitText = block.LimitBox.Text?.Trim() ?? string.Empty, InterruptCapabilities = block.InterruptCapabilities,
            Priority = block.PrioCombo.SelectedItem?.ToString() ?? "Undefined",
            Policy = block.PolicyCombo.SelectedItem?.ToString() ?? "MachineDefault", AffinityMask = block.AffinityMask,
            PowerSaving = block.PowerSavingCheck?.Checked, Wifi = block.Device.Wifi,
            NdisMode = GetSelectedNdisAffinityMode(block), RssLayout = layout, AutoMsiOnly = autoMsiOnly,
            AutoReason = block.Kind == DeviceKind.GPU && block.Device.IsIntegratedGpu ? "integrated-gpu"
                : IsAutoDisplayAudioMsiOnly(block) ? "display-audio" : "auto-policy",
        };
        return ApplyDeviceSettingsPlanAsync(block, plan, report);
    }

    private async Task ApplyDeviceSettingsPlanAsync(DeviceBlock block, DeviceSettingsWritePlan plan, OperationReport? report)
    {
        DeviceSettingsWriteResult result = await Task.Run(() =>
        {
            DeviceSettingsWriteResult outcome = new();
            try { ApplyDeviceSettingsPlan(plan, outcome); }
            catch (Exception error)
            {
                WriteLog($"APPLY.PLAN.FAILURE: {plan.InstanceId} {FlattenLogText(error.ToString())}");
                outcome.Report.AddError(plan.DeviceName, error.Message);
            }
            return outcome;
        });
        report?.MergeIssues(result.Report);
        if (_uiDisposing || IsDisposed || Disposing || block.Group.IsDisposed) return;
        if (result.NormalizedLimit is not null && block.LimitBox.Text.Trim() == plan.LimitText)
            block.LimitBox.Text = result.NormalizedLimit;
        if (result.UsbPower is not null) block.Device.UsbSelectiveSuspend = result.UsbPower;
        if (result.NicPower is not null) block.Device.NicPowerSaving = result.NicPower;
        if (result.RssRuntime is not null)
        {
            _ndisRssRuntimeCache[NormalizeInstanceId(plan.InstanceId)] = result.RssRuntime;
            block.NdisRssRuntime = result.RssRuntime;
        }
        UpdateBlockInfoText(block);
        if (report is null && result.Report.Errors.Count > 0)
            ShowOperationResult(result.Report, string.Empty, "Some device settings could not be applied.", operationName: "DEVICE SETTINGS");
    }

    internal static bool TryParseMsiLimitInput(string? input, out bool isUnlocked, out int value)
    {
        string text = input?.Trim() ?? string.Empty;
        isUnlocked = string.IsNullOrWhiteSpace(text)
            || text == "0"
            || Regex.IsMatch(text, "^(?i)(unlimited|unlocked)$", RegexOptions.CultureInvariant);
        value = 0;
        return isUnlocked
            || (Regex.IsMatch(text, "^\\d+$", RegexOptions.CultureInvariant)
                && int.TryParse(text, out value)
                && value is >= 1 and <= 2048);
    }

    internal static bool TryValidateMsiLimitInput(
        string? input,
        PciInterruptCapabilities? capabilities,
        out bool isUnlocked,
        out int value,
        out string error)
    {
        if (!TryParseMsiLimitInput(input, out isUnlocked, out value))
        {
            error = "MSI Limit must be Unlimited (or 0) or a whole number from 1 to 2048.";
            return false;
        }

        if (isUnlocked || capabilities is null)
        {
            error = string.Empty;
            return true;
        }

        PciInterruptSupport support = capabilities.Support;
        bool supportsMsi = (support & PciInterruptSupport.Msi) != 0;
        bool supportsMsiX = (support & PciInterruptSupport.MsiX) != 0;
        if (!supportsMsi && !supportsMsiX)
        {
            error = "Windows reports that the PCI device does not support MSI or MSI-X; a numeric message limit was not applied.";
            return false;
        }

        uint hardwareMaximum = capabilities.MessageMaximum is > 0
            ? Math.Min(capabilities.MessageMaximum.Value, 2048U)
            : 2048U;
        if ((uint)value > hardwareMaximum)
        {
            error = $"MSI Limit exceeds the hardware maximum reported by Windows ({hardwareMaximum}).";
            return false;
        }

        if (!supportsMsiX && supportsMsi && value is not (1 or 2 or 4 or 8 or 16))
        {
            error = "This is an MSI-only device. The allowed multi-message limits are 1, 2, 4, 8, or 16.";
            return false;
        }

        error = string.Empty;
        return true;
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
            block.AffinityLabel.Text = block.Kind == DeviceKind.STOR
                ? "Affinity Mask: Windows Default"
                : (block.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(block.Device.InstanceId, block.Device.Name) || IsDisplayAudioEndpointsText(block.Device.AudioEndpoints)))
                    ? "Affinity Mask: 0x0 (Windows Default)"
                    : (block.Kind == DeviceKind.NET_NDIS && !string.Equals(block.NdisModeCombo?.SelectedItem?.ToString(), "IRQ", StringComparison.OrdinalIgnoreCase))
                        ? "Affinity (RSS mask): 0x0"
                        : "Affinity Mask: 0x0";
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
            ResetTestBlockState(block);
            WriteLog($"RESET.TEST: {block.Device.InstanceId} kind={block.Kind} -> cleared preview priority/affinity/power");
            return;
        }

        DemandRealWrite("DEVICE SETTINGS");
        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";
        void RecordResetError(string operation, Exception ex)
        {
            report?.AddError($"{block.Device.Name}: {operation}", ex.Message);
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
            report?.AddError($"{block.Device.Name}: {operation}", ex.Message);
        }

        if (block.Kind == DeviceKind.USB)
        {
            try
            {
                UsbSelectiveSuspendPolicy.ApplyResult powerResult =
                    UsbSelectiveSuspendPolicy.ApplyControllerAndHubs(block.Device.InstanceId, enabled: true);

                block.Device.UsbSelectiveSuspend = "on";
                if (block.PowerSavingCheck is not null)
                {
                    block.PowerSavingCheck.Checked = true;
                }

                WriteLog(
                    $"RESET.POWER: {block.Device.InstanceId} -> Power Saving=Enabled " +
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
                DevicePowerPolicy.WriteStatus msPowerStatus = DevicePowerPolicy.SetDevicePowerEnable(
                    block.Device.InstanceId,
                    allowTurnOff: true,
                    out string? msPowerError);
                block.PowerSavingCheck.Checked = true;
                block.Device.NicPowerSaving = "on";
                WriteLog(
                    $"RESET.POWER: {block.Device.InstanceId} -> Power Saving=Enabled " +
                    $"PnPCapabilities=0x{pnpCaps:X} wmiWrite={msPowerStatus}");
                if (msPowerStatus is DevicePowerPolicy.WriteStatus.NotExposed or DevicePowerPolicy.WriteStatus.NotSupported)
                {
                    WriteLog($"RESET.POWER.NOTE: {block.Device.InstanceId}: MSPower_DeviceEnable status={msPowerStatus} {msPowerError}");
                }
                else if (DevicePowerPolicy.IsWriteFailure(msPowerStatus))
                {
                    WriteLog($"RESET.POWER.ERROR: {block.Device.InstanceId} wmiWrite={msPowerStatus} {msPowerError}");
                    RecordResetError(
                        "reset NIC power saving (Device Manager)",
                        new InvalidOperationException(msPowerError ?? "MSPower_DeviceEnable write failed"));
                }
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
        // forceDisable (AUTO) also requires at least one real USB block: otherwise
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

    private OperationReport SyncLivePowerManagementAfterRestore(DeviceTweakerBackup backup)
    {
        List<LivePowerRestoreTarget> targets = [];
        OperationReport report = new();
        foreach (DeviceBlock block in _blocks)
        {
            if (block.Device.IsTestDevice) continue;
            try
            {
                if (block.Kind == DeviceKind.USB)
                {
                    foreach (string id in UsbSelectiveSuspendPolicy.EnumerateBackupInstanceIds(block.Device.InstanceId))
                        targets.Add(new(id, LivePowerRestoreService.ReadRestoredEnabled(
                            backup, id, DevicePowerPolicy.TryGetClassKeyPath(id), usb: true)));
                }
                else if (block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX && !block.Device.Wifi)
                {
                    string id = block.Device.InstanceId;
                    targets.Add(new(id, LivePowerRestoreService.ReadRestoredEnabled(
                        backup, id, GetClassKeyForDevice(id), usb: false)));
                }
            }
            catch (Exception ex)
            {
                report.AddError(block.Device.InstanceId, "Restored live power targets could not be resolved.", ex.ToString());
            }
        }

        report.MergeIssues(LivePowerRestoreService.Restore(targets,
            (id, enabled) =>
            {
                DevicePowerPolicy.WriteStatus status = DevicePowerPolicy.SetDevicePowerEnable(id, enabled, out string? error);
                return new(status, error);
            },
            () =>
            {
                if (!backup.RegistryValues.Any(value => UsbSelectiveSuspendPolicy.TryParsePowerPlanTarget(
                    value.Hive, value.Path, value.Name, out _, out _))) return;
                UsbSelectiveSuspendPolicy.ActivateCurrentPowerScheme();
                WriteLog("BACKUP.RESTORE: re-activated power scheme after USB/NIC power restore");
            }, WriteLog));
        return report;
    }
}
