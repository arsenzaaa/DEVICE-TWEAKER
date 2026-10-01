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
            msiSupported = 1;
            limit = 0;
            limitPresent = true;
        }

        block.MsiCombo.SelectedItem = msiSupported == 1 ? "Enabled" : "Disabled";
        block.LimitBox.Text = limitPresent && limit > 0 ? limit.ToString() : "0";

        int prioValue = 2;
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
            3 => "High",
            _ => "Normal",
        };

        block.PolicyCombo.Items.Clear();
        if (block.Kind == DeviceKind.NET_NDIS)
        {
            block.PolicyLabel.Text = "Policy (RSS base)";
            FixRssPolicyLabelOverlap(block);
            block.PolicyCombo.Items.Add("RSS base core");
            block.PolicyCombo.SelectedIndex = 0;
            block.PolicyCombo.Enabled = false;

            int? baseCore = isTestDevice ? 0 : GetNdisBaseCore(block.Device.InstanceId);
            if (baseCore is >= 0 && baseCore < _maxLogical)
            {
                ulong mask = 1UL << baseCore.Value;
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
            }

            string loadPrefix = isTestDevice ? "LOAD.TEST" : "LOAD";
            WriteLog($"{loadPrefix}: NET_NDIS {block.Device.InstanceId} MSI={(msiSupported == 1 ? "Enabled" : "Disabled")} Limit={(limitPresent ? limit.ToString() : "Unlimited")} PrioVal={prioValue} BaseCore={(baseCore ?? -1)} Mask=0x{block.AffinityMask:X}");
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

                        object? rawOverride = affKey.GetValue("AssignmentSetOverride");
                        if (rawOverride is byte[] bytes && bytes.Length >= 4)
                        {
                            if (bytes.Length >= 8)
                            {
                                mask = BitConverter.ToUInt64(bytes, 0);
                            }
                            else
                            {
                                mask = BitConverter.ToUInt32(bytes, 0);
                            }
                        }
                        else if (rawOverride is int intVal)
                        {
                            mask = (uint)intVal;
                        }
                        else if (rawOverride is long longVal)
                        {
                            mask = (ulong)longVal;
                        }
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

                ImodConfigEntry? overrideEntry = FindImodOverride(block.Device.InstanceId, config);
                if (overrideEntry?.Enabled == false)
                {
                    block.ImodBox.Text = string.Empty;
                }
                else
                {
                    uint interval = GetEffectiveImodInterval(block.Device.InstanceId, config);
                    block.ImodBox.Text = FormatImodValue(interval);
                }
            }
            else
            {
                block.ImodBox.Text = string.Empty;
                block.ImodDefaultLabel.Text = string.Empty;
                block.ImodDefaultLabel.Tag = null;
            }
        }
        else
        {
            block.ImodBox.Text = string.Empty;
            block.ImodDefaultLabel.Text = string.Empty;
            block.ImodDefaultLabel.Tag = null;
        }

        string shortPnp = GetShortPnpId(block.Device.InstanceId);
        string displayReg = GetDisplayRegPath(block.Device.InstanceId);
        StringBuilder info = new();
        if (block.Device.IsTestDevice)
        {
            info.AppendLine("TEST DEVICE (no registry writes)");
        }
        info.AppendLine($"PNP ID: {shortPnp}");
        info.AppendLine($"Class: {block.Device.Class}");
        info.Append($"Registry: {displayReg}");

        if (block.Device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(block.Device.UsbRoles))
        {
            info.AppendLine();
            info.Append($"HID: {block.Device.UsbRoles}");
        }
        else if (block.Device.Kind == DeviceKind.NET_NDIS)
        {
            info.AppendLine();
            info.Append("Net type: NDIS (RSS)");
        }
        else if (block.Device.Kind == DeviceKind.NET_CX)
        {
            info.AppendLine();
            info.Append("Net type: NetAdapterCx");
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

        block.InfoLabel.Text = info.ToString();
        block.InfoLabel.Tag = GetFullRegPath($"HKLM\\{regBase}");
    }

    private void SaveBlockSettings(DeviceBlock block)
    {
        if (block.Device.IsTestDevice)
        {
            WriteLog($"APPLY.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason=TEST_DEVICE");
            return;
        }

        RecalcAffinityMask(block);

        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";

        try
        {
            Registry.LocalMachine.CreateSubKey(intBase)?.Dispose();
        }
        catch
        {
        }

        string msiPath = intBase + @"\MessageSignaledInterruptProperties";
        try
        {
            Registry.LocalMachine.CreateSubKey(msiPath)?.Dispose();
        }
        catch
        {
        }

        string mode = block.MsiCombo.SelectedItem?.ToString() ?? "Disabled";
        int msiVal = mode == "Enabled" ? 1 : 0;
        try
        {
            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
            msiKey?.SetValue("MSISupported", msiVal, RegistryValueKind.DWord);
        }
        catch
        {
        }

        string limitText = block.LimitBox.Text?.Trim() ?? string.Empty;
        bool isUnlimited = string.IsNullOrWhiteSpace(limitText) || limitText == "0" || Regex.IsMatch(limitText, "^(?i)unlimited$", RegexOptions.CultureInvariant);

        try
        {
            using RegistryKey? msiKey = Registry.LocalMachine.OpenSubKey(msiPath, writable: true);
            if (msiKey is not null)
            {
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
                        limitVal = 0;
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
                    MessageBox.Show(
                        "MSI Limit must be a whole number. Leave empty or set 0 for unlimited. Value has been reset to 0 (unlimited).",
                        "DEVICE TWEAKER",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    msiKey.DeleteValue("MessageNumberLimit", throwOnMissingValue: false);
                    block.LimitBox.Text = "0";
                    limitText = "0";
                }
            }
        }
        catch
        {
        }

        string prioPath = intBase + @"\Priority";
        string prioAffPath = intBase + @"\Affinity Policy";
        try
        {
            Registry.LocalMachine.CreateSubKey(prioAffPath)?.Dispose();
        }
        catch
        {
        }

        string prioStr = block.PrioCombo.SelectedItem?.ToString() ?? "Normal";
        int prioVal = prioStr switch
        {
            "Low" => 1,
            "High" => 3,
            _ => 2,
        };

        try
        {
            using RegistryKey? prioKey = Registry.LocalMachine.OpenSubKey(prioAffPath, writable: true);
            prioKey?.SetValue("DevicePriority", prioVal, RegistryValueKind.DWord);
        }
        catch
        {
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch
        {
        }

        WriteLog($"APPLY: {block.Device.InstanceId} MSI={mode} Limit={limitText} Prio={prioStr} Mask=0x{block.AffinityMask:X} Kind={block.Kind}");

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            int baseCore = 0;
            for (int i = 0; i < block.CpuBoxes.Count; i++)
            {
                if (block.CpuBoxes[i].Checked)
                {
                    baseCore = i;
                    break;
                }
            }

            WriteLog($"APPLY: NET_NDIS {block.Device.InstanceId} baseCore={baseCore}");
            SetNdisBaseCore(block.Device.InstanceId, baseCore);
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
            catch
            {
            }

            return;
        }

        try
        {
            Registry.LocalMachine.CreateSubKey(affPath)?.Dispose();
        }
        catch
        {
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
                return;
            }

            affKey.SetValue("DevicePolicy", policyVal, RegistryValueKind.DWord);
            byte[] bytes = IntPtr.Size >= 8 ? BitConverter.GetBytes(mask) : BitConverter.GetBytes((uint)mask);
            affKey.SetValue("AssignmentSetOverride", bytes, RegistryValueKind.Binary);
            WriteLog($"APPLY: AFFINITY {block.Device.InstanceId} policy={policyStr} value={policyVal} mask=0x{mask:X}");
        }
        catch
        {
        }
    }

    private void ResetBlockSettings(DeviceBlock block)
    {
        if (block.Device.IsTestDevice)
        {
            WriteLog($"RESET.SKIP: {block.Device.InstanceId} kind={block.Kind} reason=TEST_DEVICE");
            return;
        }

        string regBase = block.Device.RegBase;
        string intBase = regBase + @"\Device Parameters\Interrupt Management";

        string prioPath = intBase + @"\Priority";
        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(prioPath, throwOnMissingSubKey: false);
        }
        catch
        {
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
        catch
        {
        }

        try
        {
            Registry.LocalMachine.DeleteSubKeyTree(affPath, throwOnMissingSubKey: false);
        }
        catch
        {
        }

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            ClearNdisBaseCore(block.Device.InstanceId);
        }

        WriteLog($"RESET: {block.Device.InstanceId} kind={block.Kind} -> cleared priority/affinity (MSI left unchanged by design)");
    }
}
