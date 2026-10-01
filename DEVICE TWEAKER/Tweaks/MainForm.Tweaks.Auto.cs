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

            if (primary.EffClass > 0)
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
        bool hasE = primaryE.Count > 0;

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
                .Where(lp => lp.EffClass == 0 && (targetCcdLps.Count == 0 || targetCcdLps.Contains(lp.LP)))
                .Select(lp => lp.LP)
                .OrderBy(x => x)
                .ToList();

            allE = _cpuInfo.Topology.LPs
                .Where(lp => lp.EffClass > 0 && (targetCcdLps.Count == 0 || targetCcdLps.Contains(lp.LP)))
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

        HashSet<int> used = [];

        int? TakeLowestP(bool excludeZero)
        {
            IEnumerable<int> candidates = primaryP.Where(p => !used.Contains(p));
            if (excludeZero)
            {
                candidates = candidates.Where(p => p != 0);
            }

            int? lp = candidates.OrderBy(x => x).Cast<int?>().FirstOrDefault();
            if (lp is null)
            {
                return null;
            }

            used.Add(lp.Value);
            return lp.Value;
        }

        int? TakeHighestP(bool excludeZero)
        {
            IEnumerable<int> candidates = primaryP.Where(p => !used.Contains(p));
            if (excludeZero)
            {
                candidates = candidates.Where(p => p != 0);
            }

            int[] arr = candidates.OrderBy(x => x).ToArray();
            if (arr.Length == 0)
            {
                return null;
            }

            int lp = arr[^1];
            used.Add(lp);
            return lp;
        }

        List<int> TakeTwoAdjacentP(bool allowZero)
        {
            IEnumerable<int> candidates = primaryP.Where(p => !used.Contains(p));
            if (!allowZero)
            {
                candidates = candidates.Where(p => p != 0);
            }

            int[] arr = candidates.OrderBy(x => x).ToArray();
            if (arr.Length < 2)
            {
                return [];
            }

            for (int i = 0; i < arr.Length - 1; i++)
            {
                if (Math.Abs(arr[i + 1] - arr[i]) == 1)
                {
                    used.Add(arr[i]);
                    used.Add(arr[i + 1]);
                    return [arr[i], arr[i + 1]];
                }
            }

            used.Add(arr[0]);
            used.Add(arr[1]);
            return [arr[0], arr[1]];
        }

        int? TakeOneE()
        {
            int? lp = primaryE.Where(e => !used.Contains(e)).OrderBy(x => x).Cast<int?>().FirstOrDefault();
            if (lp is null)
            {
                return null;
            }

            used.Add(lp.Value);
            return lp.Value;
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
        List<DeviceBlock> storBlocks = _blocks.Where(b => b.Kind == DeviceKind.STOR).ToList();
        List<DeviceBlock> audioBlocks = _blocks.Where(b => b.Kind == DeviceKind.AUDIO).ToList();
        List<DeviceBlock> realAudioBlocks = [];
        HashSet<string> skipAutoIds = new(StringComparer.OrdinalIgnoreCase);
        int? audioLpForMic = null;
        bool audioLpFromSpeakers = false;

        int netCount = netNdisBlocks.Count + netCxBlocks.Count;
        bool hasWiFiOnly = wifiIds.Count > 0 && netCount == 0;
        WriteLog(
            $"AUTO.SUMMARY: GPU={gpuBlocks.Count} NET={netCount} USB={usbBlocks.Count} AUDIO={audioBlocks.Count} STOR={storBlocks.Count} WIFI={wifiIds.Count} WiFiOnly={hasWiFiOnly} targetCCD=[{string.Join(',', targetCcdLps)}] primaryP=[{string.Join(',', primaryP)}] primaryE=[{string.Join(',', primaryE)}]");
        if (hasWiFiOnly)
        {
            WriteLog("AUTO.WIFI-ONLY: no wired NET adapters found; skipping NET affinity.");
        }

        foreach (DeviceBlock audioBlock in audioBlocks)
        {
            string pnpId = audioBlock.Device.InstanceId;
            string desc = audioBlock.Device.Name;
            string audioText = audioBlock.Device.AudioEndpoints;

            bool isDisplay = IsDisplayHdmiaudio(pnpId, desc) || IsDisplayAudioEndpointsText(audioText);
            if (!isDisplay)
            {
                realAudioBlocks.Add(audioBlock);
            }
            else
            {
                skipAutoIds.Add(pnpId);
                WriteLog($"AUTO.SKIP.AUDIO: {pnpId} classified as display/HDMI audio (name=\"{desc}\" endpoints=\"{audioText}\")");
            }
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
                WriteLog($"AUTO.RESET.SKIP: {block.Device.InstanceId} Kind={block.Kind} reason=display/HDMI audio");
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

        if (realAudioBlocks.Count > 0)
        {
            foreach (DeviceBlock block in realAudioBlocks)
            {
                int? audioLp = null;
                if (hasE)
                {
                    audioLp = TakeOneE();
                }

                if (audioLp is null && Regex.IsMatch(block.Device.AudioEndpoints ?? string.Empty, "(?i)speakers"))
                {
                    if (_maxLogical > 2 && !used.Contains(2))
                    {
                        audioLp = 2;
                    }
                    else if (_maxLogical > 10 && !used.Contains(10))
                    {
                        audioLp = 10;
                    }
                }

                audioLp ??= TakeLowestP(excludeZero: true);
                audioLp ??= TakeLowestP(excludeZero: false);
                if (audioLp is null && primaryP.Count > 0)
                {
                    audioLp = primaryP[0];
                }

                if (audioLp is not null && block.Kind != DeviceKind.STOR)
                {
                    used.Add(audioLp.Value);
                    block.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in block.CpuBoxes)
                        {
                            cb.Checked = false;
                        }

                        if (audioLp.Value >= 0 && audioLp.Value < block.CpuBoxes.Count)
                        {
                            block.CpuBoxes[audioLp.Value].Checked = true;
                        }
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                    {
                        block.PolicyCombo.SelectedItem = "SpecCPU";
                    }

                    RecalcAffinityMask(block);
                    bool isSpeakers = Regex.IsMatch(block.Device.AudioEndpoints ?? string.Empty, "(?i)speakers");
                    if (audioLpForMic is null || (isSpeakers && !audioLpFromSpeakers))
                    {
                        audioLpForMic = audioLp.Value;
                        audioLpFromSpeakers = isSpeakers;
                    }
                    WriteLog($"AUTO: AUDIO {block.Device.InstanceId} -> LP={audioLp} policy={(block.PolicyCombo.SelectedItem?.ToString() ?? "(none)")}");
                }
                else
                {
                    WriteLog($"AUTO: AUDIO {block.Device.InstanceId} -> no available LP");
                }
            }
        }

        List<DeviceBlock> micUsbBlocks = usbBlocks
            .Where(b => !string.IsNullOrWhiteSpace(b.Device.UsbRoles) && Regex.IsMatch(b.Device.UsbRoles, "(?i)\\bMicrophone\\b"))
            .Where(b => !Regex.IsMatch(b.Device.UsbRoles ?? string.Empty, "(?i)\\b(Mouse|Keyboard|Gamepad)\\b"))
            .ToList();
        if (micUsbBlocks.Count > 0)
        {
            int? micLp = audioLpForMic;
            string micSource = "audio";
            if (micLp is null)
            {
                int? eLp = TakeOneE();
                if (eLp is not null)
                {
                    micLp = eLp.Value;
                    micSource = "E-core";
                }
                else
                {
                    int? pLp = TakeLowestP(excludeZero: true);
                    pLp ??= TakeLowestP(excludeZero: false);
                    if (pLp is null && primaryP.Count > 0)
                    {
                        pLp = primaryP[0];
                    }
                    if (pLp is not null)
                    {
                        micLp = pLp.Value;
                        micSource = "P-core";
                    }
                }
            }

            if (micLp is null)
            {
                foreach (DeviceBlock block in micUsbBlocks)
                {
                    WriteLog($"AUTO: USB(MIC) {block.Device.InstanceId} -> no available LP; skip");
                }
            }
            else
            {
                foreach (DeviceBlock block in micUsbBlocks)
                {
                    int lpVal = micLp.Value;
                    block.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in block.CpuBoxes)
                        {
                            cb.Checked = false;
                        }

                        if (lpVal >= 0 && lpVal < block.CpuBoxes.Count)
                        {
                            block.CpuBoxes[lpVal].Checked = true;
                        }
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                    {
                        block.PolicyCombo.SelectedItem = "SpecCPU";
                    }

                    RecalcAffinityMask(block);
                    WriteLog($"AUTO: USB(MIC) {block.Device.InstanceId} -> LP={lpVal} policy=SpecCPU source={micSource}");
                }
            }
        }

        List<DeviceBlock> hidUsbBlocks = usbBlocks
            .Where(b => !string.IsNullOrWhiteSpace(b.Device.UsbRoles)
                && Regex.IsMatch(b.Device.UsbRoles, "(?i)Mouse|Keyboard|Gamepad"))
            .ToList();
        foreach (DeviceBlock block in hidUsbBlocks)
        {
            int? lp = TakeLowestP(excludeZero: true);
            lp ??= TakeHighestP(excludeZero: true);
            lp ??= TakeLowestP(excludeZero: false);

            if (lp is not null && block.Kind != DeviceKind.STOR)
            {
                used.Add(lp.Value);
                block.SuppressCpuEvents++;
                try
                {
                    foreach (CheckBox cb in block.CpuBoxes)
                    {
                        cb.Checked = false;
                    }

                    if (lp.Value >= 0 && lp.Value < block.CpuBoxes.Count)
                    {
                        block.CpuBoxes[lp.Value].Checked = true;
                    }
                }
                finally
                {
                    block.SuppressCpuEvents--;
                }

                if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                {
                    block.PolicyCombo.SelectedItem = "SpecCPU";
                }

                RecalcAffinityMask(block);
                WriteLog($"AUTO: USB(HID) {block.Device.InstanceId} -> LP={lp} policy=SpecCPU");
            }
        }

        if (gpuBlocks.Count > 0)
        {
            List<int> gpuLps = [];
            if (pCount >= 4)
            {
                gpuLps = TakeTwoAdjacentP(allowZero: false);
            }
            else
            {
                int? gpuLp = TakeHighestP(excludeZero: true) ?? TakeHighestP(excludeZero: false);
                if (gpuLp is not null)
                {
                    gpuLps = [gpuLp.Value];
                }
            }

            foreach (DeviceBlock block in gpuBlocks)
            {
                if (gpuLps.Count > 0 && block.Kind != DeviceKind.STOR)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in block.CpuBoxes)
                        {
                            cb.Checked = false;
                        }

                        foreach (int lpVal in gpuLps)
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

                    if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                    {
                        block.PolicyCombo.SelectedItem = "SpecCPU";
                    }

                    RecalcAffinityMask(block);
                    WriteLog($"AUTO: GPU {block.Device.InstanceId} -> LPs=[{string.Join(',', gpuLps)}] policy=SpecCPU");
                }
            }
        }

        List<DeviceBlock> netBlocks = [.. netNdisBlocks, .. netCxBlocks];
        if (netBlocks.Count > 0)
        {
            int? netLp = TakeHighestP(excludeZero: true) ?? TakeLowestP(excludeZero: false);
            foreach (DeviceBlock block in netBlocks)
            {
                if (netLp is not null && block.Kind != DeviceKind.STOR)
                {
                    block.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in block.CpuBoxes)
                        {
                            cb.Checked = false;
                        }

                        if (netLp.Value >= 0 && netLp.Value < block.CpuBoxes.Count)
                        {
                            block.CpuBoxes[netLp.Value].Checked = true;
                        }
                    }
                    finally
                    {
                        block.SuppressCpuEvents--;
                    }

                    if (block.Kind != DeviceKind.NET_NDIS && block.PolicyCombo.Enabled)
                    {
                        block.PolicyCombo.SelectedItem = "SpecCPU";
                    }

                    RecalcAffinityMask(block);
                    WriteLog($"AUTO: NET {block.Device.InstanceId} -> LP={netLp} policy={block.PolicyCombo.SelectedItem}");
                }
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
