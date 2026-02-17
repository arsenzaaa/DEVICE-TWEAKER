using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private (List<int> P, List<int> E)? GetAutoCpuSets()
    {
        if (_cpuInfo is null)
        {
            return null;
        }

        List<int> primaryP = [];
        List<int> primaryE = [];

        foreach (int core in _cpuInfo.Topology.ByCore.Keys.OrderBy(x => x))
        {
            List<CpuLpInfo> group = _cpuInfo.Topology.ByCore[core];
            CpuLpInfo? primary = group.OrderBy(x => x.LP).FirstOrDefault();
            if (primary is null)
            {
                continue;
            }

            if (IsEfficiencyCore(primary))
            {
                primaryE.Add(primary.LP);
            }
            else
            {
                primaryP.Add(primary.LP);
            }
        }

        return (primaryP.OrderBy(x => x).ToList(), primaryE.OrderBy(x => x).ToList());
    }

    private void InvokeAutoOptimization()
    {
        if (_blocks.Count == 0)
        {
            return;
        }

        WriteLog("AUTO: Invoke-AutoOptimization start");
        if (_testAutoDryRun)
        {
            ResetReservedCpuSetsPreview();
        }
        else
        {
            ResetReservedCpuSets();
        }

        (List<int> P, List<int> E)? cpuSets = GetAutoCpuSets();
        if (cpuSets is null)
        {
            return;
        }

        List<int> primaryP = cpuSets.Value.P.Where(lp => lp >= 0 && lp < _maxLogical).ToList();
        List<int> primaryE = cpuSets.Value.E.Where(lp => lp >= 0 && lp < _maxLogical).ToList();
        List<int> targetCcdLps = [];
        List<int> ccdIdsUnique = [];
        if (_cpuInfo is not null && _cpuInfo.CcdMap.Count > 0)
        {
            ccdIdsUnique = _cpuInfo.CcdMap.Values.Distinct().OrderBy(x => x).ToList();
            if (ccdIdsUnique.Count >= 2)
            {
                int targetCcd = ccdIdsUnique[^1];
                targetCcdLps = _cpuInfo.CcdMap.Where(kvp => kvp.Value == targetCcd).Select(kvp => kvp.Key).OrderBy(x => x).ToList();
                primaryP = primaryP.Where(lp => targetCcdLps.Contains(lp)).ToList();
                primaryE = primaryE.Where(lp => targetCcdLps.Contains(lp)).ToList();

                WriteLog(
                    $"AUTO: CCD map ids=[{string.Join(',', ccdIdsUnique)}] targetCCD={targetCcd} targetLPs=[{string.Join(',', targetCcdLps)}] filteredP=[{string.Join(',', primaryP)}] filteredE=[{string.Join(',', primaryE)}]");

                if (primaryP.Count + primaryE.Count == 0 && targetCcdLps.Count > 0)
                {
                    primaryP = targetCcdLps.OrderBy(x => x).ToList();
                    WriteLog($"AUTO: CCD fallback -> using all target CCD LPs as primary P: [{string.Join(',', primaryP)}]");
                }
            }
        }

        List<int> allP = [];
        List<int> allE = [];
        if (_cpuInfo is not null)
        {
            allP = _cpuInfo.Topology.LPs
                .Where(lp => !IsEfficiencyCore(lp) && (targetCcdLps.Count == 0 || targetCcdLps.Contains(lp.LP)))
                .Select(lp => lp.LP)
                .OrderBy(x => x)
                .ToList();

            allE = _cpuInfo.Topology.LPs
                .Where(lp => IsEfficiencyCore(lp) && (targetCcdLps.Count == 0 || targetCcdLps.Contains(lp.LP)))
                .Select(lp => lp.LP)
                .OrderBy(x => x)
                .ToList();
        }

        if (primaryP.Count == 0 && allP.Count > 0)
        {
            primaryP = allP;
            WriteLog($"AUTO: P fallback -> using all P cores from map: [{string.Join(',', primaryP)}]");
        }

        if (primaryE.Count == 0 && allE.Count > 0)
        {
            primaryE = allE;
        }

        int pCount = primaryP.Count;
        if (pCount <= 0)
        {
            return;
        }

        WriteLog($"AUTO: CPU primary P=[{string.Join(',', primaryP)}] E=[{string.Join(',', primaryE)}]");

        List<int> coreOrder = primaryP.Count > 0 ? primaryP : primaryE;
        bool usingP = primaryP.Count > 0;
        bool zeroAvailable = coreOrder.Contains(0);
        List<int> available = [];
        int coreIndex = 0;
        int extras = 0;
        Queue<int> audioQueue = new(primaryE.Where(lp => lp != 0));
        int audioECount = 0;
        int audioEAssigned = 0;
        bool useAudioE = primaryP.Count > 0 && audioQueue.Count > 0;

        List<int> TakeNext(int count)
        {
            List<int> lps = [];
            while (lps.Count < count && coreIndex < available.Count)
            {
                lps.Add(available[coreIndex]);
                coreIndex++;
            }

            if (lps.Count == count && extras > 0 && coreIndex < available.Count)
            {
                coreIndex++;
                extras--;
            }

            return lps;
        }

        foreach (IGrouping<DeviceKind, DeviceBlock> g in _blocks.GroupBy(b => b.Kind))
        {
            WriteLog($"AUTO: blocks kind={g.Key} count={g.Count()}");
        }

        HashSet<string> wifiIds = new(StringComparer.OrdinalIgnoreCase);
        foreach (DeviceBlock nb in _blocks.Where(b => b.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX))
        {
            if (nb.Device.Wifi)
            {
                wifiIds.Add(nb.Device.InstanceId);
                WriteLog($"AUTO.WIFI.DETECT: {nb.Device.InstanceId} name=\"{nb.Device.Name}\"");
            }
        }

        List<DeviceBlock> usbBlocks = _blocks.Where(b => b.Kind == DeviceKind.USB).ToList();
        List<DeviceBlock> netNdisBlocks = _blocks.Where(b => b.Kind == DeviceKind.NET_NDIS && !wifiIds.Contains(b.Device.InstanceId)).ToList();
        List<DeviceBlock> netCxBlocks = _blocks.Where(b => b.Kind == DeviceKind.NET_CX && !wifiIds.Contains(b.Device.InstanceId)).ToList();
        List<DeviceBlock> gpuBlocks = _blocks.Where(b => b.Kind == DeviceKind.GPU).ToList();
        List<DeviceBlock> audioBlocks = _blocks.Where(b => b.Kind == DeviceKind.AUDIO).ToList();
        int storCount = _blocks.Count(b => b.Kind == DeviceKind.STOR);
        HashSet<string> skipAutoIds = new(StringComparer.OrdinalIgnoreCase);

        foreach (DeviceBlock usbBlock in usbBlocks)
        {
            if (string.IsNullOrWhiteSpace(usbBlock.Device.UsbRoles))
            {
                skipAutoIds.Add(usbBlock.Device.InstanceId);
                WriteLog($"AUTO.SKIP.USB: {usbBlock.Device.InstanceId} no HID roles (manual/reset only)");
            }
        }

        int netCount = netNdisBlocks.Count + netCxBlocks.Count;
        bool hasWiFiOnly = wifiIds.Count > 0 && netCount == 0;
        WriteLog(
            $"AUTO.SUMMARY: GPU={gpuBlocks.Count} NET={netCount} USB={usbBlocks.Count} AUDIO={audioBlocks.Count} STOR={storCount} WIFI={wifiIds.Count} WiFiOnly={hasWiFiOnly} targetCCD=[{string.Join(',', targetCcdLps)}] primaryP=[{string.Join(',', primaryP)}] primaryE=[{string.Join(',', primaryE)}]");
        if (hasWiFiOnly)
        {
            WriteLog("AUTO.WIFI-ONLY: no wired NET adapters found; skipping NET affinity.");
        }

        foreach (DeviceBlock audioBlock in audioBlocks)
        {
            string pnpId = audioBlock.Device.InstanceId;
            string desc = audioBlock.Device.Name;
            string audioText = audioBlock.Device.AudioEndpoints;

            bool isSpdif = IsSpdifAudioEndpointsText(audioText);
            bool isDisplay = IsDisplayHdmiaudio(pnpId, desc) || IsDisplayAudioEndpointsText(audioText);
            if (!isDisplay && !isSpdif)
            {
                continue;
            }

            skipAutoIds.Add(pnpId);
            string reason = isSpdif ? "digital S/PDIF audio" : "display/HDMI audio";
            WriteLog($"AUTO.SKIP.AUDIO: {pnpId} classified as {reason} (name=\"{desc}\" endpoints=\"{audioText}\")");
        }

        foreach (DeviceBlock block in _blocks)
        {
            bool isSkipAuto = skipAutoIds.Contains(block.Device.InstanceId);
            bool isWifi = wifiIds.Contains(block.Device.InstanceId);
            ulong beforeMask = block.AffinityMask;
            string beforePolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog($"AUTO.RESET: {block.Device.InstanceId} Kind={block.Kind} maskBefore=0x{beforeMask:X} policyBefore={beforePolicy}");

            if (isWifi)
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                WriteLog($"AUTO.WIFI.SKIP: {block.Device.InstanceId} -> MSI=Enabled Prio=High (affinity/limit preserved)");
                continue;
            }

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

            if (isSkipAuto)
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "Normal";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }
            else
            {
                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "MachineDefault";
                }
            }

            RecalcAffinityMask(block);

            ulong afterMask = block.AffinityMask;
            string afterPolicy = block.PolicyCombo.SelectedItem?.ToString() ?? "(none)";
            WriteLog($"AUTO.RESET: {block.Device.InstanceId} Kind={block.Kind} maskAfter=0x{afterMask:X} policyAfter={afterPolicy} skipAuto={isSkipAuto}");
            if (isSkipAuto)
            {
                string reason = IsSpdifAudioEndpointsText(block.Device.AudioEndpoints) ? "digital S/PDIF audio" : "display/HDMI audio";
                WriteLog($"AUTO.RESET.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason={reason}");
            }
        }

        bool anyChecked = usbBlocks.Any(b => IsUsbImodTarget(b.Device) && b.ImodAutoCheck.Checked);
        if (!anyChecked)
        {
            WriteLog("AUTO.IMOD: no checkboxes selected -> skipping");
        }
        else
        {
            string imodDefault = FormatImodValue(ImodDefaultInterval);
            string imodDisabled = FormatImodValue(0);
            foreach (DeviceBlock block in usbBlocks.Where(b => IsUsbImodTarget(b.Device)))
            {
                if (!block.ImodAutoCheck.Checked)
                {
                    WriteLog($"AUTO.IMOD.SKIP: {block.Device.InstanceId} checkbox=off");
                    continue;
                }

                string roles = block.Device.UsbRoles ?? string.Empty;
                bool hasKeyboard = Regex.IsMatch(roles, "(?i)\\bKeyboard\\b");
                bool hasMouse = Regex.IsMatch(roles, "(?i)\\bMouse\\b");

                string before = block.ImodBox.Text?.Trim() ?? string.Empty;
                if (hasKeyboard && hasMouse)
                {
                    block.ImodBox.Text = imodDisabled;
                    WriteLog($"AUTO.IMOD: {block.Device.InstanceId} -> {imodDisabled} (roles=\"{roles}\", prev={before})");
                }
                else
                {
                    block.ImodBox.Text = imodDefault;
                    WriteLog($"AUTO.IMOD: {block.Device.InstanceId} -> {imodDefault} (roles=\"{roles}\", prev={before})");
                }
            }
        }

        List<DeviceBlock> orderedBlocks = _blocks
            .Where(b => b.Kind != DeviceKind.STOR)
            .Where(b => !wifiIds.Contains(b.Device.InstanceId))
            .Where(b => !skipAutoIds.Contains(b.Device.InstanceId))
            .ToList();

        int audioCount = orderedBlocks.Count(b => b.Kind == DeviceKind.AUDIO);
        if (useAudioE && audioCount > 0)
        {
            audioECount = Math.Min(audioCount, audioQueue.Count);
        }

        int totalNeeded = orderedBlocks.Sum(b => b.Kind == DeviceKind.GPU ? 2 : 1) - audioECount;
        if (totalNeeded < 0)
        {
            totalNeeded = 0;
        }
        available = coreOrder.Where(lp => lp != 0).ToList();
        if (available.Count < totalNeeded && zeroAvailable)
        {
            available.Add(0);
        }

        extras = Math.Max(0, available.Count - totalNeeded);
        coreIndex = 0;

        foreach (DeviceBlock block in orderedBlocks)
        {
            int need = block.Kind == DeviceKind.GPU ? 2 : 1;
            List<int> lps;
            if (useAudioE && block.Kind == DeviceKind.AUDIO && audioEAssigned < audioECount)
            {
                lps = audioQueue.Count > 0 ? [audioQueue.Dequeue()] : [];
                if (lps.Count > 0)
                {
                    audioEAssigned++;
                }
            }
            else
            {
                lps = TakeNext(need);
            }

            if (lps.Count == 0)
            {
                string label = block.Kind == DeviceKind.NET_NDIS || block.Kind == DeviceKind.NET_CX ? "NET" : block.Kind.ToString();
                WriteLog($"AUTO: {label} {block.Device.InstanceId} -> no available LP ({(usingP ? "P" : "E")}-core list exhausted)");
                continue;
            }

            block.SuppressCpuEvents++;
            try
            {
                foreach (CheckBox cb in block.CpuBoxes)
                {
                    cb.Checked = false;
                }

                foreach (int lpVal in lps)
                {
                    if (lpVal >= 0 && lpVal < block.CpuBoxes.Count)
                    {
                        block.CpuBoxes[lpVal].Checked = true;
                    }
                }
            }
            finally
            {
                block.SuppressCpuEvents--;
            }

            int? ndisAutoQueues = null;
            if (block.Kind == DeviceKind.NET_NDIS)
            {
                // AUTO keeps NDIS on a single logical processor to avoid implicit SMT spread via RSS queues.
                block.RssBaseCore = lps[0];
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

                    ndisAutoQueues = (int)block.RssQueueBox.Value;
                }
            }

            if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
            {
                block.PolicyCombo.SelectedItem = "SpecCPU";
            }

            RecalcAffinityMask(block);

            string labelText = block.Kind switch
            {
                DeviceKind.NET_NDIS or DeviceKind.NET_CX => "NET",
                _ => block.Kind.ToString(),
            };
            if (block.Kind == DeviceKind.NET_NDIS)
            {
                WriteLog($"AUTO: {labelText} {block.Device.InstanceId} -> LPs=[{string.Join(',', lps)}] policy={block.PolicyCombo.SelectedItem} queues={(ndisAutoQueues ?? 1)}");
            }
            else
            {
                WriteLog($"AUTO: {labelText} {block.Device.InstanceId} -> LPs=[{string.Join(',', lps)}] policy={block.PolicyCombo.SelectedItem}");
            }
        }

        WriteLog("AUTO: Invoke-AutoOptimization done");
    }

    private void ResetAllTweaks()
    {
        WriteLog("RESET: full reset requested");
        if (_blocks.Count == 0)
        {
            RefreshBlocks();
        }

        foreach (DeviceBlock b in _blocks)
        {
            try
            {
                ResetBlockSettings(b);
                LoadBlockSettings(b);
            }
            catch (Exception ex)
            {
                WriteLog($"RESET.ERROR: {b.Device.InstanceId} -> {ex.Message}");
            }
        }

        ResetReservedCpuSets();
        ResetImodIntervalsToDefault();

        LogGuiSnapshot("reset-all");
        ShowThemedInfo(
            "All Device Tweaker changes have been cleared.\nPlease reboot your PC to fully revert device behavior.");
    }

    private void ResetImodIntervalsToDefault()
    {
        string defaultText = FormatImodValue(ImodDefaultInterval);
        bool hasUsb = false;

        foreach (DeviceBlock block in _blocks)
        {
            if (IsUsbImodTarget(block.Device))
            {
                string before = block.ImodBox.Text?.Trim() ?? string.Empty;
                block.ImodBox.Text = defaultText;
                block.ImodAutoCheck.Checked = false;
                WriteLog($"RESET.IMOD.USB: {block.Device.InstanceId} {before} -> {defaultText}");
                hasUsb = true;
            }
        }

        if (hasUsb)
        {
            _ = ApplyImodSettings(out string? note);
            if (!string.IsNullOrWhiteSpace(note))
            {
                WriteLog($"RESET.IMOD.USB: {note}");
            }
        }

        RemoveImodPersistenceFiles();
    }
}
