namespace DeviceTweakerCS;

// Self-contained parts of a device card; NewDeviceBlock assembles and lays them out.
public sealed partial class MainForm
{
    private DeviceCard NewDeviceCard(DeviceInfo device, string title)
    {
        (string badge, Color color) = device.Kind switch
        {
            DeviceKind.USB => ("[USB]", DeviceBadgePalette.Usb),
            DeviceKind.GPU => GetGpuKindBadge(device),
            DeviceKind.NET_NDIS or DeviceKind.NET_CX => ("[NET]", DeviceBadgePalette.Network),
            DeviceKind.AUDIO => ("[AUDIO]", DeviceBadgePalette.Audio),
            DeviceKind.STOR => GetStorageKindBadge(device),
            _ => ("[DEV]", _statusPrefix),
        };
        DeviceCard card = new(new(_uiScale, _blockFont, _blockTitleFont, _bgGroup, _fgMain, _mutedText, _border), title, badge, color);
        if (device.Kind == DeviceKind.USB && device.UsbChipPath is UsbChipPathInfo chip)
            _copyToolTip.SetToolTip(card.TitleLabel, chip.TooltipText);
        return card;
    }

    // One checkbox per logical CPU, grouped into columns by CCD (AMD) or P/E cores
    // (Intel hybrid). Returns the content extent the card layout needs.
    private (List<CheckBox> Boxes, int ContentRight, int RequiredWidth, int ContentHeight) BuildAffinityCheckboxes(
        DeviceInfo device, BufferedPanel cpuPanel)
    {
        List<CheckBox> cpuBoxes = [];
        List<(int Lp, CheckBox Control, int Ccd, int Eff)> lpMeta = [];
        int checkSpacing = UiScale(22);

        for (int i = 0; i < _maxLogical; i++)
        {
            CheckBox cb = new()
            {
                Text = $"CPU {i}",
                Font = _blockFont,
                AutoSize = true,
                ForeColor = _fgMain,
                BackColor = _bgForm,
                FlatStyle = FlatStyle.Flat,
            };
            StyleCpuCheckbox(cb, i);
            cb.Tag = i;
            if (device.Kind == DeviceKind.STOR)
            {
                cb.AutoCheck = false;
                cb.TabStop = false;
                cb.Cursor = Cursors.No;
                cb.ForeColor = _border;
            }

            cpuBoxes.Add(cb);

            int ccdId = _cpuInfo?.CcdMap.TryGetValue(i, out int cid) == true ? cid : 0;
            int eff = _cpuLpByIndex.TryGetValue(i, out CpuLpInfo? lpInfoLocal) ? lpInfoLocal.EffClass : -1;
            lpMeta.Add((i, cb, ccdId, eff));
            cpuPanel.Controls.Add(cb);
        }

        List<int> ccdKeys = lpMeta.Select(m => m.Ccd).Distinct().OrderBy(x => x).ToList();
        if (ccdKeys.Count == 0)
        {
            ccdKeys.Add(0);
        }

        List<List<(int Lp, CheckBox Control, int Ccd, int Eff)>> columns = [];
        foreach (int cid in ccdKeys)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> items = lpMeta.Where(m => m.Ccd == cid).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> pItems = items.Where(m => !IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> eItems = items.Where(m => IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> other = items.Except(pItems).Except(eItems).ToList();
            if (ccdKeys.Count == 1 && pItems.Count > 0 && eItems.Count > 0)
            {
                // Intel Hybrid: place P-Cores and E-Cores into dedicated side-by-side columns
                columns.Add(pItems);
                columns.Add([.. eItems, .. other]);
            }
            else if (ccdKeys.Count == 1 && items.Count > 16)
            {
                // Large single-CCD: split into two balanced columns to prevent clipping
                int mid = (items.Count + 1) / 2;
                columns.Add(items.Take(mid).ToList());
                columns.Add(items.Skip(mid).ToList());
            }
            else
            {
                List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = [.. pItems, .. eItems, .. other];
                columns.Add(ordered);
            }
        }

        int columnGap = UiScale(16);
        int startX = UiScale(10);
        int minColumnWidth = UiScale(160);
        // Text-only MeasureText misses the native glyph; add the system check size
        // plus a small DPI gutter instead of a huge fixed pad that forces H-scroll.
        int checkboxGlyphGutter = SystemInformation.MenuCheckSize.Width + UiScale(20);

        int runningX = startX;
        int maxColumnCount = 0;
        List<int> columnWidths = [];
        foreach (List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered in columns)
        {
            if (ordered.Count > maxColumnCount)
            {
                maxColumnCount = ordered.Count;
            }

            int maxWidth = minColumnWidth;
            if (ordered.Count > 0)
            {
                int w = ordered.Max(o =>
                {
                    // GetPreferredSize for a Standard CheckBox is only reliable
                    // after the native handle exists. Without this, long labels
                    // such as CPPC rank plus CCD/CCX are measured too narrowly
                    // and lose their trailing token when painted.
                    o.Control.CreateControl();
                    Size measured = TextRenderer.MeasureText(
                        o.Control.Text,
                        o.Control.Font,
                        new Size(int.MaxValue, int.MaxValue),
                        TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
                    int preferred = o.Control.GetPreferredSize(Size.Empty).Width;
                    return Math.Max(preferred, measured.Width + checkboxGlyphGutter);
                });
                if (w > 0)
                {
                    // A standard WinForms CheckBox can discard the entire trailing
                    // core-type token when its client width is only a few pixels
                    // below the native preferred width. Two-digit CPU labels expose
                    // that native rendering behavior, so retain a measured gutter.
                    maxWidth = Math.Max(minColumnWidth, w + UiScale(12));
                }
            }

            columnWidths.Add(maxWidth);
        }

        if (columnWidths.Count > 1)
        {
            int uniformColumnWidth = columnWidths.Max();
            for (int i = 0; i < columnWidths.Count; i++)
            {
                columnWidths[i] = uniformColumnWidth;
            }
        }

        for (int i = 0; i < columns.Count; i++)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = columns[i];
            int maxWidth = columnWidths[i];
            int cellWidth = Math.Max(minColumnWidth, maxWidth);
            int cellHeight = Math.Max(UiScale(18), checkSpacing - UiScale(2));
            int y = UiScale(4);
            foreach ((int _, CheckBox control, int _, int _) in ordered)
            {
                control.AutoSize = false;
                control.Size = new Size(cellWidth, cellHeight);
                control.TextAlign = ContentAlignment.MiddleLeft;
                control.Location = new Point(runningX, y);
                y += checkSpacing;
            }

            runningX += maxWidth + columnGap;
        }

        int cpuContentRight = columns.Count == 0 ? 0 : runningX - columnGap;
        int requiredWidth = columns.Count == 0 ? cpuPanel.Width : cpuContentRight + startX + UiScale(12);
        int cpuContentHeight = Math.Max((maxColumnCount * checkSpacing) + UiScale(18), UiScale(150));
        return (cpuBoxes, cpuContentRight, requiredWidth, cpuContentHeight);
    }

    // Flat card button with the same hover style as the top bar.
    private Button NewCardButton(string resourceKey, Size size)
    {
        Button button = new()
        {
            Text = UiResources.English(resourceKey),
            Size = size,
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(button);
        button.MouseEnter += (_, _) => SetTopButtonHoverStyle(button);
        button.MouseLeave += (_, _) => SetTopButtonBaseStyle(button);
        return button;
    }

    // CPU checkboxes, RSS, NDIS mode, mouse throttle and NIC ITR actions of one card.
    private void WireDeviceBlockControls(DeviceBlock block, List<CheckBox> cpuBoxes)
    {
    foreach (CheckBox cb in cpuBoxes)
    {
        cb.CheckedChanged += (_, _) =>
        {
            if (block.SuppressCpuEvents == 0)
            {
                if (block.Kind == DeviceKind.NET_NDIS)
                {
                    HandleNdisCheckboxChanged(block, cb);
                }
                else
                {
                    RecalcAffinityMask(block);
                }
                OnBlockSettingChanged(block);
            }
        };
    }

    if (block.Kind == DeviceKind.NET_NDIS && block.RssQueueBox is not null)
    {
        block.RssQueueBox.ValueChanged += (_, _) =>
        {
            if (block.SuppressCpuEvents > 0)
            {
                return;
            }

            int baseCore = block.RssBaseCore ?? GetFirstCheckedCore(block) ?? 0;
            int requested = ClampRssQueueCount((int)block.RssQueueBox.Value);
            int queues = block.RssQueueOptions is { Count: > 0 } options
                ? NdisRssLayout.Snap(options, block.RssQueueLast, requested)
                : requested;
            ApplyNdisSelection(block, baseCore, queues);
        };
    }

    if (block.NdisModeCombo is not null)
    {
        block.NdisModeCombo.SelectedIndexChanged += (_, _) =>
        {
            RecalcAffinityMask(block);
            if (block.SuppressCpuEvents == 0) UpdateBlockInfoText(block);
        };
    }

    if (block.RawMouseThrottleCheck is not null && block.RawMouseThrottleCombo is not null)
    {
        block.RawMouseThrottleCheck.CheckedChanged += (_, _) => HandleRawMouseThrottleChanged(block);
        block.RawMouseThrottleCombo.SelectedIndexChanged += (_, _) => HandleRawMouseThrottleChanged(block);
    }

    if (block.NicItrApplyButton is not null)
    {
        block.NicItrApplyButton.Click += async (_, _) => await ApplyNicItrFromBlock(block);
    }

    if (block.NicItrSaveButton is not null)
    {
        block.NicItrSaveButton.Click += async (_, _) => await SaveNicItrPersistenceFromBlock(block);
    }

    if (block.NicItrCheckButton is not null)
    {
        block.NicItrCheckButton.Click += async (_, _) => await CheckImodDriverFromNicBlock(block);
    }

    if (block.NicItrBox is not null)
    {
        block.NicItrBox.TextChanged += (_, _) =>
        {
            if (block.NicItrEditorUpdateDepth == 0) block.NicItrEditGeneration++;
            UpdateNicItrInputTimeLabel(block);
        };
    }
    }

    // Every editable control marks the card as changed (APPLY counter, "Changed" badge).
    private void WireDeviceBlockChangeTracking(DeviceBlock block, bool hasImod)
    {
    block.MsiCombo.SelectedIndexChanged += (_, _) => OnBlockSettingChanged(block);
    if (block.LimitBox is not null)
    {
        block.LimitBox.TextChanged += (_, _) => OnBlockSettingChanged(block);
    }
    block.PrioCombo.SelectedIndexChanged += (_, _) => OnBlockSettingChanged(block);
    block.PolicyCombo.SelectedIndexChanged += (_, _) => OnBlockSettingChanged(block);
    if (block.PowerSavingCheck is not null)
    {
        block.PowerSavingCheck.CheckedChanged += (_, _) => OnBlockSettingChanged(block);
    }
    if (block.NdisModeCombo is not null)
    {
        block.NdisModeCombo.SelectedIndexChanged += (_, _) => OnBlockSettingChanged(block);
    }
    if (block.RssQueueBox is not null)
    {
        block.RssQueueBox.ValueChanged += (_, _) => OnBlockSettingChanged(block);
    }
    if (block.RawMouseThrottleCheck is not null)
    {
        block.RawMouseThrottleCheck.CheckedChanged += (_, _) => OnBlockSettingChanged(block);
    }
    if (block.RawMouseThrottleCombo is not null)
    {
        block.RawMouseThrottleCombo.SelectedIndexChanged += (_, _) => OnBlockSettingChanged(block);
    }
    if (hasImod)
    {
        block.ImodAutoCheck.CheckedChanged += (_, _) => OnBlockSettingChanged(block);
        if (block.ImodBox is not null)
        {
            block.ImodBox.TextChanged += (_, _) => OnBlockSettingChanged(block);
        }
    }
    }
}
