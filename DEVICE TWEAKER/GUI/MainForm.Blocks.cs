using System.Diagnostics;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const string StorageAffinityNoteText = "(affinity masks not supported on SSD/HDD)";
    private static readonly string[] ImodDeviceEditorRoleOrder = ["Mouse", "Keyboard", "Audio", "Gamepad"];

    private static IReadOnlyList<string> GetImodDeviceEditorRoles(DeviceInfo device)
    {
        HashSet<string> roles = new(StringComparer.OrdinalIgnoreCase);
        AddImodDeviceEditorRoles(device.UsbRoles, roles);
        AddImodDeviceEditorRoles(device.AudioEndpoints, roles);

        return ImodDeviceEditorRoleOrder
            .Where(roles.Contains)
            .ToArray();
    }

    private static void AddImodDeviceEditorRoles(string? text, HashSet<string> roles)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        foreach (string part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (TryGetAdaptiveRoleBase(part, out string role)
                && ImodDeviceEditorRoleOrder.Contains(role, StringComparer.OrdinalIgnoreCase))
            {
                roles.Add(role);
            }
        }
    }

    private static uint GetImodDeviceEditorDefaultValue(string role)
    {
        return role.Equals("Mouse", StringComparison.OrdinalIgnoreCase)
            ? 0x0
            : role.Equals("Audio", StringComparison.OrdinalIgnoreCase)
                ? 0xFA0
                : ImodDefaultInterval;
    }

    private int GetDevicesScrollGutterWidth()
    {
        if (_devicesScroll is null)
        {
            return 0;
        }

        return _devicesScroll.Width + UiScale(2);
    }

    private int GetDevicesViewportWidth()
    {
        int width = _devicesHost.ClientSize.Width;
        if (width <= 0)
        {
            width = _devicesPanel.ClientSize.Width;
        }

        width -= GetDevicesScrollGutterWidth();

        return Math.Max(UiScale(360) + (UiScale(24) * 2), width);
    }

    private void FixRssPolicyLabelOverlap(DeviceBlock block)
    {
        int desiredLeft = block.PolicyLabel.Right + UiScale(8);
        if (desiredLeft <= block.PolicyCombo.Left)
        {
            return;
        }

        int rightPadding = UiScale(24);
        int parentLeft = block.PolicyCombo.Parent?.Left ?? 0;
        int maxWidth = block.Group.ClientSize.Width - rightPadding - parentLeft - desiredLeft;
        int newWidth = block.PolicyCombo.Width;
        if (maxWidth < newWidth)
        {
            newWidth = maxWidth;
        }

        if (newWidth < UiScale(90))
        {
            newWidth = UiScale(90);
        }

        block.PolicyCombo.Location = new Point(desiredLeft, block.PolicyCombo.Top);
        block.PolicyCombo.Width = newWidth;
    }

    private string BuildDeviceBlockTitle(DeviceInfo device, string? usbRolesOverride = null)
    {
        string title = device.Name;
        string usbRoles = usbRolesOverride ?? device.UsbRoles;
        if (device.Kind == DeviceKind.GPU && device.IsIntegratedGpu)
        {
            title = $"{device.Name} [iGPU]";
        }
        else if (device.Kind == DeviceKind.USB && !string.IsNullOrWhiteSpace(usbRoles))
        {
            string chip = device.UsbChipPath is { } path ? $" | {path.CompactTag}" : string.Empty;
            title = $"{device.Name} [{usbRoles}{chip}]";
        }
        else if (device.Kind == DeviceKind.USB)
        {
            title = device.UsbChipPath is { } path
                ? $"{device.Name} [No HID roles | {path.CompactTag}]"
                : $"{device.Name} [No HID roles]";
        }
        else if (device.Kind == DeviceKind.NET_NDIS)
        {
            title = $"{device.Name} [NDIS]";
        }
        else if (device.Kind == DeviceKind.NET_CX)
        {
            title = $"{device.Name} [NetAdapterCx]";
        }
        else if (device.Kind == DeviceKind.STOR)
        {
            if (string.Equals(device.StorageTag, "SSD+HDD", StringComparison.OrdinalIgnoreCase))
            {
                title = $"{device.Name} [SSD+HDD]";
            }
            else if (!string.IsNullOrWhiteSpace(device.StorageTag) &&
                     !string.Equals(device.StorageTag, "NVMe", StringComparison.OrdinalIgnoreCase) &&
                     !string.Equals(device.StorageTag, "SSD", StringComparison.OrdinalIgnoreCase) &&
                     !string.Equals(device.StorageTag, "HDD", StringComparison.OrdinalIgnoreCase))
            {
                title = $"{device.Name} [{device.StorageTag}]";
            }
        }
        else if (device.Kind == DeviceKind.AUDIO && !string.IsNullOrWhiteSpace(device.AudioEndpoints))
        {
            title = $"{device.Name} [{device.AudioEndpoints}]";
        }

        if (device.IsTestDevice)
        {
            title = $"[TEST] {title}";
        }

        return title;
    }

    private static (string kindText, Color kindColor) GetStorageKindBadge(DeviceInfo device)
    {
        bool isNvme = Regex.IsMatch(device.Name ?? string.Empty, "(?i)NVM\\s*Express|NVMe")
            || string.Equals(device.StorageTag, "NVMe", StringComparison.OrdinalIgnoreCase);
        if (isNvme)
        {
            return ("[NVMe]", Color.FromArgb(140, 200, 255));
        }

        if (string.Equals(device.StorageTag, "SSD", StringComparison.OrdinalIgnoreCase))
        {
            return ("[SSD]", Color.FromArgb(120, 220, 200));
        }

        if (string.Equals(device.StorageTag, "HDD", StringComparison.OrdinalIgnoreCase))
        {
            return ("[HDD]", Color.FromArgb(180, 180, 190));
        }

        if (Regex.IsMatch(device.Name ?? string.Empty, "(?i)\\bSATA\\b|\\bAHCI\\b"))
        {
            return ("[SATA]", Color.FromArgb(160, 200, 220));
        }

        return ("[STORAGE]", Color.FromArgb(190, 190, 200));
    }

    private void NewDeviceBlock(
        DeviceInfo device,
        int index,
        IReadOnlyDictionary<string, string>? priorImodStatuses = null)
    {
        DeviceBlock? createdBlock = null;
        DeviceCardPanel grp = new()
        {
            BorderColor = _border,
        };

        string title = BuildDeviceBlockTitle(device);
        string logTitle = device.Kind == DeviceKind.STOR ? $"{title} {StorageAffinityNoteText}" : title;

        WriteLog(
            $"UI.BLOCK: idx={index} title=\"{logTitle}\" kind={device.Kind} id={device.InstanceId} roles=\"{device.UsbRoles}\" audio=\"{device.AudioEndpoints}\" storage=\"{device.StorageTag}\"");

        grp.Width = GetDevicesViewportWidth() - UiScale(40);
        grp.Height = _grpHeight;
        grp.BackColor = _bgGroup;
        grp.ForeColor = _fgMain;
        grp.Font = _blockFont;
        grp.Margin = new Padding(0);
        grp.Padding = new Padding(UiScale(12), UiScale(16), UiScale(12), UiScale(16));

        FlowLayoutPanel headerPanel = new()
        {
            AutoSize = false,
            Size = new Size(grp.Width - UiScale(40), UiScale(24)),
            Location = new Point(UiScale(18), UiScale(8)),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = Color.Transparent,
        };

        (string kindText, Color kindColor) = device.Kind switch
        {
            DeviceKind.USB => ("[USB]", Color.FromArgb(120, 210, 255)),
            DeviceKind.GPU => ("[GPU]", Color.FromArgb(195, 150, 255)),
            DeviceKind.NET_NDIS or DeviceKind.NET_CX => ("[NET]", Color.FromArgb(100, 225, 155)),
            DeviceKind.AUDIO => ("[AUDIO]", Color.FromArgb(255, 155, 125)),
            DeviceKind.STOR => GetStorageKindBadge(device),
            _ => ("[DEV]", _statusPrefix)
        };

        Label kindBadge = new()
        {
            Text = kindText,
            Font = _blockTitleFont,
            ForeColor = kindColor,
            AutoSize = true,
            Margin = new Padding(0, 0, UiScale(6), 0),
        };
        headerPanel.Controls.Add(kindBadge);

        Label headerLabel = new HighlightLabel
        {
            Text = title,
            Font = _blockTitleFont,
            ForeColor = _fgMain,
            AutoSize = true,
            Margin = Padding.Empty,
            HighlightText = "Mouse scanning",
            HighlightColor = Color.FromArgb(120, 120, 120),
        };
        headerPanel.Controls.Add(headerLabel);
        if (device.Kind == DeviceKind.USB && device.UsbChipPath is UsbChipPathInfo chipTip)
        {
            _copyToolTip.SetToolTip(headerLabel, chipTip.TooltipText);
        }

        Label? headerNote = null;

        Label modifiedBadge = new()
        {
            Text = UiLanguage.Text("[ MODIFIED ]"),
            Font = _blockTitleFont,
            ForeColor = _statusWarn,
            AutoSize = true,
            Visible = false,
            Margin = new Padding(UiScale(8), 0, 0, 0),
        };
        headerPanel.Controls.Add(modifiedBadge);

        Panel divider = new()
        {
            BackColor = _border,
            Size = new Size(grp.Width - UiScale(32), UiScale(1)),
            Location = new Point(UiScale(16), UiScale(36)),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };

        int contentTop = UiScale(48);
        Label cpuLabel = new()
        {
            Text = "CPU Affinity",
            AutoSize = true,
            ForeColor = _fgMain,
            Location = new Point(UiScale(18), contentTop),
        };
        if (device.Kind == DeviceKind.STOR)
        {
            cpuLabel.ForeColor = _mutedText;
        }

        int cpuPanelTop = cpuLabel.Bottom + UiScale(6);
        int cpuPanelHeight = UiScale(150);
        int settingsMinimumWidth = Math.Min(UiScale(420), Math.Max(UiScale(320), grp.Width - UiScale(48)));
        int desiredSettingsSideWidth = UiScale(600);
        int settingsSideMinimumWidth = desiredSettingsSideWidth;
        int cpuPanelMinimumWidth = UiScale(320);
        int settingsSideGap = UiScale(40);
        int cpuPanelFullMaximumWidth = Math.Max(
            cpuPanelMinimumWidth,
            grp.Width - UiScale(16) - UiScale(24));
        int cpuPanelSideMaximumWidth = Math.Max(
            cpuPanelMinimumWidth,
            grp.Width - UiScale(16) - settingsSideGap - settingsMinimumWidth - UiScale(24));
        int cpuPanelMaximumWidth = cpuPanelSideMaximumWidth;
        int cpuPanelWidth = cpuPanelMinimumWidth;
        // Dual-CCD affinity lists are wide; a plain Panel + AutoScroll paints overlapping
        // checkbox columns. BufferedPanel keeps the scroll viewport clean.
        bool affinitySettingsStacked = false;
        BufferedPanel cpuPanel = new()
        {
            Location = new Point(UiScale(16), cpuPanelTop),
            Size = new Size(cpuPanelWidth, cpuPanelHeight),
            BackColor = _bgForm,
            Font = _blockFont,
            Padding = new Padding(UiScale(8), UiScale(6), UiScale(8), UiScale(6)),
        };
        cpuPanel.Paint += (_, e) =>
        {
            Rectangle rect = cpuPanel.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };

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

        FlowLayoutPanel? affinityQuickPanel = null;
        if (device.Kind != DeviceKind.STOR && device.Kind != DeviceKind.NET_NDIS)
        {
            affinityQuickPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Location = new Point(cpuLabel.Right + UiScale(12), contentTop - UiScale(2)),
                Margin = Padding.Empty,
                Padding = Padding.Empty,
            };

            Button NewAffinityQuickBtn(string text, string tip, Action onClick)
            {
                Button btn = new ThemedButton
                {
                    Text = text,
                    Font = _blockFont,
                    Height = UiScale(20),
                    AutoSize = true,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = _bgGroup,
                    ForeColor = _statusInactive,
                    Cursor = Cursors.Hand,
                    TabStop = false,
                    Margin = new Padding(UiScale(4), 0, 0, 0),
                };
                btn.FlatAppearance.BorderSize = 1;
                btn.FlatAppearance.BorderColor = Color.FromArgb(65, 65, 75);
                btn.MouseEnter += (_, _) =>
                {
                    btn.ForeColor = _fgMain;
                    btn.FlatAppearance.BorderColor = _accent;
                };
                btn.MouseLeave += (_, _) =>
                {
                    btn.ForeColor = _statusInactive;
                    btn.FlatAppearance.BorderColor = Color.FromArgb(65, 65, 75);
                };
                btn.Click += (_, _) => onClick();
                _copyToolTip.SetToolTip(btn, tip);
                return btn;
            }

            affinityQuickPanel.Controls.Add(NewAffinityQuickBtn(
                UiLanguage.Text("[ ALL ]"),
                UiLanguage.Text("Select all logical cores"),
                () =>
                {
                    if (createdBlock is null) return;
                    createdBlock.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in cpuBoxes) cb.Checked = true;
                    }
                    finally
                    {
                        createdBlock.SuppressCpuEvents--;
                    }
                    RecalcAffinityMask(createdBlock);
                    OnBlockSettingChanged(createdBlock);
                    WriteLog($"UI: Affinity quick-select ALL device={device.InstanceId} mask=0x{createdBlock.AffinityMask:X}");
                }));

            affinityQuickPanel.Controls.Add(NewAffinityQuickBtn(
                UiLanguage.Text("[ NONE ]"),
                UiLanguage.Text("Clear all cores"),
                () =>
                {
                    if (createdBlock is null) return;
                    createdBlock.SuppressCpuEvents++;
                    try
                    {
                        foreach (CheckBox cb in cpuBoxes) cb.Checked = false;
                    }
                    finally
                    {
                        createdBlock.SuppressCpuEvents--;
                    }
                    RecalcAffinityMask(createdBlock);
                    OnBlockSettingChanged(createdBlock);
                    WriteLog($"UI: Affinity quick-select NONE device={device.InstanceId} mask=0x{createdBlock.AffinityMask:X}");
                }));

            if (_effClassP.Count > 0)
            {
                affinityQuickPanel.Controls.Add(NewAffinityQuickBtn(
                    UiLanguage.Text("[ P-CORES ]"),
                    UiLanguage.Text("Select physical P-cores only"),
                    () =>
                    {
                        if (createdBlock is null) return;
                        createdBlock.SuppressCpuEvents++;
                        try
                        {
                            foreach ((int lp, CheckBox cb, _, int eff) in lpMeta)
                            {
                                cb.Checked = _effClassP.Contains(eff);
                            }
                        }
                        finally
                        {
                            createdBlock.SuppressCpuEvents--;
                        }
                        RecalcAffinityMask(createdBlock);
                        OnBlockSettingChanged(createdBlock);
                        WriteLog($"UI: Affinity quick-select P-CORES device={device.InstanceId} mask=0x{createdBlock.AffinityMask:X}");
                    }));
            }

            if (ccdKeys.Count > 1)
            {
                affinityQuickPanel.Controls.Add(NewAffinityQuickBtn(
                    "[ CCD0 ]",
                    UiLanguage.Text("Select CCD0 cores only"),
                    () =>
                    {
                        if (createdBlock is null) return;
                        createdBlock.SuppressCpuEvents++;
                        try
                        {
                            foreach ((int _, CheckBox cb, int ccd, _) in lpMeta)
                            {
                                cb.Checked = (ccd == 0);
                            }
                        }
                        finally
                        {
                            createdBlock.SuppressCpuEvents--;
                        }
                        RecalcAffinityMask(createdBlock);
                        OnBlockSettingChanged(createdBlock);
                        WriteLog($"UI: Affinity quick-select CCD0 device={device.InstanceId} mask=0x{createdBlock.AffinityMask:X}");
                    }));
            }
        }

        List<List<(int Lp, CheckBox Control, int Ccd, int Eff)>> columns = [];
        foreach (int cid in ccdKeys)
        {
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> items = lpMeta.Where(m => m.Ccd == cid).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> pItems = items.Where(m => !IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> eItems = items.Where(m => IsEfficiencyClass(m.Eff)).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> other = items.Except(pItems).Except(eItems).ToList();
            List<(int Lp, CheckBox Control, int Ccd, int Eff)> ordered = [.. pItems, .. eItems, .. other];
            columns.Add(ordered);
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
                    maxWidth = Math.Max(minColumnWidth, w + UiScale(28));
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
        int requiredWidth = columns.Count == 0 ? cpuPanel.Width : cpuContentRight + startX + UiScale(18);
        int cpuPanelHorizontalSlack = UiScale(30);
        int cpuContentHeight = Math.Max((maxColumnCount * checkSpacing) + UiScale(18), UiScale(150));
        int sideSettingsX = 0;
        int sideSettingsWidth = 0;
        int settingsX = UiScale(18);
        int availableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
        void UpdateResponsivePlacement()
        {
            settingsSideMinimumWidth = Math.Min(desiredSettingsSideWidth, Math.Max(UiScale(440), grp.Width - UiScale(48)));
            settingsMinimumWidth = Math.Min(UiScale(420), Math.Max(UiScale(320), grp.Width - UiScale(48)));
            cpuPanelFullMaximumWidth = Math.Max(cpuPanelMinimumWidth, grp.Width - cpuPanel.Left - UiScale(24));
            cpuPanelSideMaximumWidth = Math.Max(
                cpuPanelMinimumWidth,
                grp.Width - cpuPanel.Left - settingsSideGap - settingsSideMinimumWidth - UiScale(24));

            int neededCpuWidth = Math.Max(cpuPanelMinimumWidth, requiredWidth + cpuPanelHorizontalSlack);
            // Dual-CCD / long CPPC labels: if two columns cannot sit beside MSI/IRQ
            // settings without a horizontal scrollbar, stack settings under the list.
            affinitySettingsStacked = neededCpuWidth > cpuPanelSideMaximumWidth + UiScale(4);
            cpuPanelMaximumWidth = affinitySettingsStacked
                ? cpuPanelFullMaximumWidth
                : cpuPanelSideMaximumWidth;
            int targetCpuPanelWidth = Math.Min(cpuPanelMaximumWidth, neededCpuWidth);
            if (cpuPanel.Width != targetCpuPanelWidth)
            {
                cpuPanel.Width = targetCpuPanelWidth;
            }

            if (cpuPanel.Height != cpuContentHeight)
            {
                cpuPanel.Height = cpuContentHeight;
            }

            bool needsHorizontalScroll = cpuContentRight > cpuPanel.ClientSize.Width + UiScale(2);
            if (needsHorizontalScroll)
            {
                cpuPanel.AutoScroll = true;
                cpuPanel.AutoScrollMinSize = new Size(
                    requiredWidth + cpuPanelHorizontalSlack,
                    Math.Max(cpuContentHeight, cpuPanel.ClientSize.Height));
            }
            else
            {
                cpuPanel.AutoScroll = false;
                cpuPanel.AutoScrollMinSize = Size.Empty;
                cpuPanel.HorizontalScroll.Visible = false;
                cpuPanel.VerticalScroll.Visible = false;
            }

            if (affinitySettingsStacked)
            {
                settingsX = cpuPanel.Left;
                sideSettingsX = settingsX;
                sideSettingsWidth = cpuPanelFullMaximumWidth;
                availableSettingsWidth = Math.Max(UiScale(260), cpuPanelFullMaximumWidth);
            }
            else
            {
                sideSettingsX = cpuPanel.Right + settingsSideGap;
                sideSettingsWidth = grp.Width - sideSettingsX - UiScale(24);
                settingsX = sideSettingsX;
                availableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
            }
        }

        UpdateResponsivePlacement();

        int maskY = cpuPanel.Bottom + UiScale(10);
        string defaultMaskText = device.Kind == DeviceKind.STOR
            ? "Affinity Mask: Windows Default"
            : (device.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(device.InstanceId, device.Name) || IsDisplayAudioEndpointsText(device.AudioEndpoints)))
                ? "Affinity Mask: 0x0 (Windows Default)"
                : "Affinity Mask: 0x0";

        Label lblMask = new()
        {
            Text = defaultMaskText,
            Font = _blockFont,
            AutoSize = true,
            ForeColor = _accent,
            Location = new Point(UiScale(18), maskY),
        };
        if (device.Kind == DeviceKind.STOR)
        {
            lblMask.ForeColor = _mutedText;
            _copyToolTip.SetToolTip(lblMask, "NVMe/SATA storage uses Windows Multi-Queue steering. Pinned core affinity is intentionally disabled to ensure maximum SSD speed and low latency.");
        }
        else if (device.Kind == DeviceKind.AUDIO && (IsDisplayHdmiaudio(device.InstanceId, device.Name) || IsDisplayAudioEndpointsText(device.AudioEndpoints)))
        {
            lblMask.ForeColor = _mutedText;
            _copyToolTip.SetToolTip(lblMask, "Display/HDMI audio shares PCIe bus with the GPU and uses Windows default interrupt steering (0x0).");
        }
        else
        {
            _copyToolTip.SetToolTip(lblMask, "Interrupt Affinity Mask (AssignmentSetOverride). Strictly routes hardware interrupt service routines (ISRs) and deferred procedure calls (DPCs) to the selected CPU logical cores.");
        }

        Label lblIrq = new()
        {
            Text = "IRQ Count: reading...",
            Font = _blockFont,
            AutoSize = true,
            ForeColor = _mutedText,
            Location = new Point(UiScale(18), maskY + UiScale(20)),
        };
        _copyToolTip.SetToolTip(lblIrq, "Shows the number of allocated interrupt vectors and current interrupt delivery mode. MSI/MSI-X indicates a modern dedicated interrupt vector. Line-based indicates legacy INTx sharing with other PCI devices.");

        // Longest common left labels: "Mouse Throttle:", "IRQ Priority:", "Power Saving:".
        int valueX = UiScale(140);
        int rowGap = UiScale(12);
        int rowTop = 0;
        int labelOffset = UiScale(4);
        bool showPowerSaving = device.Kind == DeviceKind.USB
            || ((device.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX) && !device.Wifi);

        Panel settingsPanel = new()
        {
            AutoSize = false,
            BackColor = Color.Transparent,
        };

        Label lblMsi = new()
        {
            Text = "MSI Mode:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbMsi = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(150, 26),
        };
        StyleDarkDropDownPicker(cmbMsi);
        cmbMsi.Items.AddRange(new object[] { "Disabled", "Enabled" });
        cmbMsi.DropDownWidth = cmbMsi.Width;
        cmbMsi.MaxDropDownItems = 2;

        const string msiTip = "Message Signaled Interrupts (MSI/MSI-X). Replaces legacy pin-based line IRQs with direct in-band PCIe memory writes to the local APIC. Eliminates interrupt sharing, lowers latency to sub-microsecond levels, and prevents DPC spikes in games.";
        _copyToolTip.SetToolTip(lblMsi, msiTip);
        _copyToolTip.SetToolTip(cmbMsi, msiTip);

        settingsPanel.Controls.AddRange([lblMsi, cmbMsi]);
        rowTop = cmbMsi.Bottom + rowGap;

        Label lblLimit = new()
        {
            Text = "MSI Limit:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedTextBox txtLimit = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(100, 24),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Center,
            Text = "0",
            Font = _blockFont,
        };
        StyleDarkTextBox(txtLimit, leftMargin: 4, rightMargin: 4);

        Label lblLimitHint = new()
        {
            Text = "(0 = unlocked)",
            AutoSize = true,
            ForeColor = _mutedText,
            Location = new Point(txtLimit.Right + UiScale(8), txtLimit.Top + UiScale(4)),
        };

        const string limitTip = "MessageNumberLimit (registry: MessageSignaledInterruptProperties). Controls the maximum number of MSI-X interrupt vectors the device driver can allocate. 0 = unlimited (hardware default). Recommended: 0 for GPUs and modern NICs.";
        _copyToolTip.SetToolTip(lblLimit, limitTip);
        _copyToolTip.SetToolTip(txtLimit, limitTip);
        _copyToolTip.SetToolTip(lblLimitHint, limitTip);

        settingsPanel.Controls.AddRange([lblLimit, txtLimit, lblLimitHint]);
        rowTop = txtLimit.Bottom + rowGap;

        Label lblPrio = new()
        {
            Text = "IRQ Priority:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbPrio = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(150, 26),
        };
        StyleDarkDropDownPicker(cmbPrio);
        cmbPrio.Items.AddRange(new object[] { "Undefined", "Low", "Normal", "High" });
        cmbPrio.DropDownWidth = cmbPrio.Width;
        cmbPrio.MaxDropDownItems = 4;

        const string prioTip = "DevicePriority (registry: Affinity Policy). Controls Windows kernel interrupt servicing priority relative to other hardware devices. Setting High ensures that critical gaming inputs (mouse, keyboard) and GPU interrupts are processed ahead of secondary devices during heavy CPU load.";
        _copyToolTip.SetToolTip(lblPrio, prioTip);
        _copyToolTip.SetToolTip(cmbPrio, prioTip);

        settingsPanel.Controls.AddRange([lblPrio, cmbPrio]);
        rowTop = cmbPrio.Bottom + rowGap;

        Label lblPolicy = new()
        {
            Text = "Policy:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
        };

        ThemedDropDownPicker cmbPolicy = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(170, 26),
        };
        StyleDarkDropDownPicker(cmbPolicy);
        cmbPolicy.DropDownWidth = cmbPolicy.Width;
        cmbPolicy.MaxDropDownItems = 6;
        if (device.Kind == DeviceKind.STOR)
        {
            lblPolicy.ForeColor = _mutedText;
            cmbPolicy.Enabled = false;
        }

        const string policyTip = "DevicePolicy (registry: Affinity Policy). Defines how the Windows HAL routes device interrupts across processors:\n• SpecifiedProcessors: strictly binds interrupts to the selected Affinity Mask\n• MachineDefault: default Windows steering via BIOS/ACPI tables\n• AllCloseProcessors / OneCloseProcessor: routes to near NUMA node cores\n• SpreadMessagesAcrossAllProcessors: distributes MSI-X messages across all cores.";
        _copyToolTip.SetToolTip(lblPolicy, policyTip);
        _copyToolTip.SetToolTip(cmbPolicy, policyTip);

        if (device.Kind != DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblPolicy, cmbPolicy]);
            rowTop = cmbPolicy.Bottom + rowGap;
        }
        else
        {
            lblPolicy.Visible = false;
            cmbPolicy.Visible = false;
        }

        Label lblNdisMode = new()
        {
            Text = "NDIS Mode:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };

        ThemedDropDownPicker cmbNdisMode = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(94, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };
        cmbNdisMode.Items.AddRange(["RSS", "IRQ", "BOTH"]);
        cmbNdisMode.SelectedIndex = 0;
        cmbNdisMode.DropDownWidth = cmbNdisMode.Width;
        cmbNdisMode.MaxDropDownItems = 3;

        if (device.Kind == DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblNdisMode, cmbNdisMode]);
            _copyToolTip.SetToolTip(cmbNdisMode, "RSS assigns receive queues to CPUs. IRQ writes interrupt-affinity policy. BOTH writes both settings when the adapter and driver support RSS.");
            rowTop = cmbNdisMode.Bottom + rowGap;
        }

        Label lblRssQueues = new()
        {
            Text = "RSS Queues:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = device.Kind == DeviceKind.NET_NDIS,
        };

        ThemedNumericUpDown nudRssQueues = new()
        {
            Location = new Point(valueX, rowTop),
            Size = UiScale(70, 26),
            Minimum = 1,
            Maximum = Math.Max(1, _maxLogical),
            Value = 1,
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            BorderStyle = BorderStyle.FixedSingle,
            Visible = device.Kind == DeviceKind.NET_NDIS,
            ButtonBackColor = Color.FromArgb(14, 14, 17),
            ButtonHoverColor = Color.FromArgb(40, 40, 48),
            ArrowColor = _fgMain,
        };

        if (device.Kind == DeviceKind.NET_NDIS)
        {
            settingsPanel.Controls.AddRange([lblRssQueues, nudRssQueues]);
            _copyToolTip.SetToolTip(nudRssQueues, "Number of RSS receive queues to configure. The adapter and driver determine the supported range. The plan starts at the selected base CPU.");
            rowTop = nudRssQueues.Bottom + rowGap;
        }

        // After interrupt/policy (and NDIS/RSS): Device Manager power-saving toggle.
        ThemedCheckBox? chkPowerSaving = null;
        if (showPowerSaving)
        {
            Label lblPowerSaving = new()
            {
                Text = "Power Saving:",
                AutoSize = true,
                Location = new Point(0, rowTop + labelOffset),
            };

            chkPowerSaving = new ThemedCheckBox
            {
                Text = "Enabled",
                SyncCheckedStateText = true,
                AutoSize = false,
                Location = new Point(valueX, rowTop + UiScale(2)),
                Size = UiScale(118, 22),
                BackColor = _bgGroup,
                ForeColor = _fgMain,
                Cursor = Cursors.Hand,
                Checked = true,
                BorderColor = _border,
                BoxBackColor = _bgGroup,
                CheckedBackColor = Color.FromArgb(12, 12, 15),
                HoverBackColor = Color.FromArgb(22, 22, 26),
                PressedBackColor = Color.FromArgb(34, 34, 40),
                CheckColor = _fgMain,
            };

            settingsPanel.Controls.AddRange([lblPowerSaving, chkPowerSaving]);
            rowTop = Math.Max(lblPowerSaving.Bottom, chkPowerSaving.Bottom) + rowGap;

            string powerTip = device.Kind == DeviceKind.USB
                ? "Same idea as Device Manager → Power Management → 'Allow the computer to turn off this device to save power'.\n" +
                  "For USB this disables Selective Suspend on the controller + root hubs and the power-plan USB SS setting.\n" +
                  "Unchecked = Disabled. Applied with APPLY / AUTO. A reboot may be required."
                : "Device Manager → Power Management → 'Allow the computer to turn off this device to save power'.\n" +
                  "Unchecked sets PnPCapabilities bit 0x08 (do-not-turn-off) and clears MSPower_DeviceEnable on the wired NIC.\n" +
                  "Applied with APPLY / AUTO. A reboot may be required.";
            _copyToolTip.SetToolTip(lblPowerSaving, powerTip);
            _copyToolTip.SetToolTip(chkPowerSaving, powerTip);
        }

        NicItrProfile? nicItrProfile = device.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX
            ? TryGetNicItrProfile(device.InstanceId)
            : null;
        bool showNicItr = nicItrProfile is not null;
        int nicSetButtonWidth = UiScale(50);
        int nicSaveButtonWidth = UiScale(58);
        int nicButtonGap = UiScale(6);
        int nicInlineGap = UiScale(8);
        int nicInputAvailableWidth = Math.Max(UiScale(100), availableSettingsWidth - valueX - UiScale(8));
        int nicInputMinBase = nicItrProfile is { MaxQueues: > 1 } ? UiScale(260) : UiScale(130);
        int nicInputMinWidth = Math.Min(nicInputMinBase, nicInputAvailableWidth);
        int nicInputDesiredWidth = nicItrProfile is { MaxQueues: > 1 } ? UiScale(330) : UiScale(160);
        int nicInlineMaxWidth = availableSettingsWidth - valueX - nicInlineGap - nicSetButtonWidth - nicButtonGap - nicSaveButtonWidth - UiScale(8);
        bool nicButtonsInline = nicInlineMaxWidth >= nicInputMinWidth;
        int nicInputWidth = nicButtonsInline
            ? Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInlineMaxWidth))
            : Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInputAvailableWidth));
        int readbackInset = UiScale(14);
        int nicReadbackX = valueX + readbackInset;
        int nicStatusWidth = Math.Max(UiScale(120), availableSettingsWidth - nicReadbackX - UiScale(8));
        int nicDetailRows = nicItrProfile is { MaxQueues: > 1 and <= 4 } ? nicItrProfile.MaxQueues : 2;
        int nicDetailHeight = Math.Max(UiScale(42), UiScale(((nicDetailRows + 1) * 17) + 8));
        Label lblNicItr = new()
        {
            Text = "NIC ITR:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = _fgMain,
            Visible = showNicItr,
        };

        ThemedTextBox txtNicItr = new()
        {
            Location = new Point(valueX, rowTop),
            Size = new Size(nicInputWidth, UiScale(24)),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Left,
            Text = "0x0",
            Visible = showNicItr,
            Font = _blockFont,
        };
        StyleDarkTextBox(txtNicItr);

        Button btnNicItr = new()
        {
            Text = "SET",
            Size = new Size(nicSetButtonWidth, UiScale(24)),
            Location = nicButtonsInline
                ? new Point(txtNicItr.Right + nicInlineGap, txtNicItr.Top)
                : new Point(valueX, txtNicItr.Bottom + UiScale(6)),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            Visible = showNicItr,
        };
        SetTopButtonBaseStyle(btnNicItr);
        btnNicItr.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnNicItr);
        btnNicItr.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnNicItr);

        Button btnNicItrSave = new()
        {
            Text = "SAVE",
            Size = new Size(nicSaveButtonWidth, UiScale(24)),
            Location = new Point(btnNicItr.Right + nicButtonGap, btnNicItr.Top),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            Visible = showNicItr,
        };
        SetTopButtonBaseStyle(btnNicItrSave);
        btnNicItrSave.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnNicItrSave);
        btnNicItrSave.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnNicItrSave);

        int nicCheckButtonWidth = UiScale(76);
        Button btnNicItrCheck = new()
        {
            Text = "CHECK",
            Size = new Size(nicCheckButtonWidth, UiScale(24)),
            Location = new Point(btnNicItrSave.Right + nicButtonGap, btnNicItr.Top),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
            Visible = showNicItr,
        };
        SetTopButtonBaseStyle(btnNicItrCheck);
        btnNicItrCheck.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnNicItrCheck);
        btnNicItrCheck.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnNicItrCheck);

        // Recalc inline layout to fit SET+SAVE+CHECK on one row when possible.
        int nicButtonsRowWidth = nicSetButtonWidth + nicButtonGap + nicSaveButtonWidth + nicButtonGap + nicCheckButtonWidth;
        nicInlineMaxWidth = availableSettingsWidth - valueX - nicInlineGap - nicButtonsRowWidth - UiScale(8);
        nicButtonsInline = nicInlineMaxWidth >= nicInputMinWidth;
        nicInputWidth = nicButtonsInline
            ? Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInlineMaxWidth))
            : Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, nicInputAvailableWidth));
        txtNicItr.Size = new Size(nicInputWidth, UiScale(24));
        if (nicButtonsInline)
        {
            btnNicItr.Location = new Point(txtNicItr.Right + nicInlineGap, txtNicItr.Top);
            btnNicItrSave.Location = new Point(btnNicItr.Right + nicButtonGap, btnNicItr.Top);
            btnNicItrCheck.Location = new Point(btnNicItrSave.Right + nicButtonGap, btnNicItr.Top);
        }
        else
        {
            btnNicItr.Location = new Point(valueX, txtNicItr.Bottom + UiScale(6));
            btnNicItrSave.Location = new Point(btnNicItr.Right + nicButtonGap, btnNicItr.Top);
            btnNicItrCheck.Location = new Point(btnNicItrSave.Right + nicButtonGap, btnNicItr.Top);
        }

        Label lblNicItrStatus = new HighlightLabel()
        {
            Text = "current: reading...",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = false,
            Size = new Size(nicStatusWidth, UiScale(22)),
            ForeColor = _statusInactive,
            Location = new Point(nicReadbackX, Math.Max(txtNicItr.Bottom, btnNicItrCheck.Bottom) + UiScale(4)),
            Visible = showNicItr,
            UseMnemonic = false,
        };

        Label lblNicItrTime = new NicItrTableLabel()
        {
            Text = "time: reading...",
            AutoSize = false,
            Size = new Size(nicStatusWidth, nicDetailHeight),
            ForeColor = _statusInactive,
            BackColor = _bgGroup,
            Font = _technicalFont,
            Location = new Point(nicReadbackX, lblNicItrStatus.Bottom + UiScale(2)),
            Visible = showNicItr,
            UseMnemonic = false,
        };

        void LayoutNicItrRow()
        {
            if (!showNicItr)
            {
                return;
            }

            int usableSettingsWidth = Math.Max(UiScale(120), availableSettingsWidth - UiScale(4));
            int nicValueX = Math.Max(valueX, lblNicItr.PreferredWidth + UiScale(14));
            nicReadbackX = nicValueX + readbackInset;
            txtNicItr.Location = new Point(nicValueX, txtNicItr.Top);

            int buttonsWidth = btnNicItr.Width + nicButtonGap
                + btnNicItrSave.Width + nicButtonGap
                + btnNicItrCheck.Width;
            int inputAvailable = Math.Max(UiScale(100), usableSettingsWidth - nicValueX - UiScale(8));
            int inlineAvailable = usableSettingsWidth - nicValueX - nicInlineGap - buttonsWidth - UiScale(8);
            bool inline = inlineAvailable >= nicInputMinWidth;
            int inputWidth = inline
                ? Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, inlineAvailable))
                : Math.Min(nicInputDesiredWidth, Math.Max(nicInputMinWidth, inputAvailable));

            txtNicItr.Size = new Size(inputWidth, UiScale(24));
            btnNicItr.Location = inline
                ? new Point(txtNicItr.Right + nicInlineGap, txtNicItr.Top)
                : new Point(nicValueX, txtNicItr.Bottom + UiScale(6));
            btnNicItrSave.Location = new Point(btnNicItr.Right + nicButtonGap, btnNicItr.Top);
            btnNicItrCheck.Location = new Point(btnNicItrSave.Right + nicButtonGap, btnNicItr.Top);

            int statusWidth = Math.Max(UiScale(120), usableSettingsWidth - nicReadbackX - UiScale(8));
            lblNicItrStatus.Location = new Point(
                nicReadbackX,
                Math.Max(txtNicItr.Bottom, btnNicItrCheck.Bottom) + UiScale(4));
            lblNicItrStatus.Size = new Size(statusWidth, UiScale(22));
            lblNicItrTime.Location = new Point(nicReadbackX, lblNicItrStatus.Bottom + UiScale(2));
            lblNicItrTime.Size = new Size(statusWidth, nicDetailHeight);
        }

        LayoutNicItrRow();

        if (showNicItr)
        {
            settingsPanel.Controls.AddRange([lblNicItr, txtNicItr, btnNicItr, btnNicItrSave, btnNicItrCheck, lblNicItrStatus, lblNicItrTime]);
            rowTop = lblNicItrTime.Bottom + rowGap;
            _copyToolTip.SetToolTip(btnNicItrCheck, "Load DTIMOD.sys (IMOD driver) and re-read current NIC ITR values.");
        }

        bool showRawMouseThrottle = HasMouseThrottleContext(device);
        bool rawMouseThrottleInteractive = showRawMouseThrottle && SupportsRawMouseThrottleOs();
        Label lblRawMouseThrottle = new()
        {
            Text = "Mouse Throttle:",
            AutoSize = true,
            Location = new Point(0, rowTop + labelOffset),
            ForeColor = rawMouseThrottleInteractive ? _fgMain : _mutedText,
            Visible = showRawMouseThrottle,
        };

        ThemedCheckBox chkRawMouseThrottle = new()
        {
            Text = "Enabled",
            SyncCheckedStateText = true,
            AutoSize = false,
            Location = new Point(valueX, rowTop + UiScale(2)),
            Size = UiScale(118, 22),
            BackColor = _bgGroup,
            ForeColor = rawMouseThrottleInteractive ? _fgMain : _mutedText,
            Cursor = rawMouseThrottleInteractive ? Cursors.Hand : Cursors.Default,
            Enabled = rawMouseThrottleInteractive,
            Visible = showRawMouseThrottle,
            BorderColor = _border,
            BoxBackColor = _bgGroup,
            CheckedBackColor = Color.FromArgb(12, 12, 15),
            HoverBackColor = Color.FromArgb(22, 22, 26),
            PressedBackColor = Color.FromArgb(34, 34, 40),
            CheckColor = _fgMain,
        };

        ThemedDropDownPicker cmbRawMouseThrottle = new()
        {
            Location = new Point(chkRawMouseThrottle.Right + UiScale(10), rowTop),
            Size = UiScale(86, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = rawMouseThrottleInteractive ? _fgMain : _mutedText,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
            Enabled = rawMouseThrottleInteractive,
            Visible = showRawMouseThrottle,
        };
        cmbRawMouseThrottle.DropDownWidth = cmbRawMouseThrottle.Width;
        cmbRawMouseThrottle.MaxDropDownItems = 7;

        Label lblRawMouseThrottleStatus = new HighlightLabel()
        {
            Text = rawMouseThrottleInteractive ? "current: off" : "current: unavailable (Windows 11 22H2+)",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = true,
            ForeColor = _statusInactive,
            Location = new Point(cmbRawMouseThrottle.Right + UiScale(8), rowTop + labelOffset),
            Visible = showRawMouseThrottle,
        };

        if (showRawMouseThrottle)
        {
            settingsPanel.Controls.AddRange([lblRawMouseThrottle, chkRawMouseThrottle, cmbRawMouseThrottle, lblRawMouseThrottleStatus]);
            rowTop = cmbRawMouseThrottle.Bottom + rowGap;
            string throttleTip = GetRawMouseThrottleToolTip(rawMouseThrottleInteractive);
            _copyToolTip.SetToolTip(lblRawMouseThrottle, throttleTip);
            _copyToolTip.SetToolTip(chkRawMouseThrottle, throttleTip);
            _copyToolTip.SetToolTip(cmbRawMouseThrottle, throttleTip);
            _copyToolTip.SetToolTip(lblRawMouseThrottleStatus, throttleTip);
        }

        int imodCheckSize = UiScale(14);
        int imodCheckGap = UiScale(4);

        CheckBox chkImod = new()
        {
            AutoSize = false,
            Size = new Size(imodCheckSize, imodCheckSize),
            Checked = false,
            BackColor = _bgGroup,
            ForeColor = _fgMain,
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false,
        };

        Label lblImod = new()
        {
            Text = "IMOD Value:",
            AutoSize = true,
            ForeColor = _fgMain,
        };

        Label lblImodMode = new()
        {
            Text = "Mode:",
            AutoSize = true,
            ForeColor = _fgMain,
        };

        ThemedDropDownPicker cmbImodMode = new()
        {
            Size = UiScale(156, 26),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            Font = _blockFont,
            BorderColor = _border,
            ButtonColor = Color.FromArgb(14, 14, 17),
            SelectedBackColor = Color.FromArgb(48, 48, 58),
            SelectedForeColor = _fgMain,
            ArrowColor = _fgMain,
            ItemHeight = UiScale(18),
        };
        cmbImodMode.Items.AddRange([ImodModeSingle, ImodModeVector, ImodModeRoles]);
        cmbImodMode.SelectedIndex = 0;
        cmbImodMode.DropDownWidth = cmbImodMode.Width;
        cmbImodMode.MaxDropDownItems = 3;

        Label lblImodModeHint = new()
        {
            Text = "single value for controller",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _statusInactive,
        };

        ThemedTextBox txtImod = new()
        {
            Size = UiScale(360, 24),
            BackColor = Color.FromArgb(18, 18, 22),
            ForeColor = _fgMain,
            TextAlign = HorizontalAlignment.Left,
            Text = "0x0",
            Font = _blockFont,
        };
        StyleDarkTextBox(txtImod);

        Label lblImodDefault = new()
        {
            Text = "default: 0x0",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _mutedText,
            Cursor = Cursors.Hand,
        };
        lblImodDefault.Click += (_, _) =>
        {
            if (lblImodDefault.Tag is string value && !string.IsNullOrWhiteSpace(value))
            {
                txtImod.Text = value;
            }
        };

        string? priorImodCurrent = null;
        if (priorImodStatuses is not null
            && !string.IsNullOrWhiteSpace(device.InstanceId)
            && priorImodStatuses.TryGetValue(NormalizeInstanceId(device.InstanceId), out string? cachedImodStatus)
            && !string.IsNullOrWhiteSpace(cachedImodStatus)
            && !cachedImodStatus.Equals("current: reading...", StringComparison.OrdinalIgnoreCase))
        {
            priorImodCurrent = cachedImodStatus;
        }

        Label lblImodCurrent = new HighlightLabel()
        {
            Text = priorImodCurrent ?? "current: -",
            HighlightText = "current:",
            HighlightColor = _statusPrefix,
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _statusInactive,
        };

        ImodMapTextBox lblImodMap = new()
        {
            Text = "devices: -",
            Font = _technicalFont,
            BackColor = _bgGroup,
            ForeColor = _mutedText,
            PrefixColor = _statusPrefix,
            RoleColor = _statusActive,
            ValueColor = _fgMain,
            TabStop = false,
            Cursor = Cursors.IBeam,
        };

        lblImodMap.DoubleClick += (_, _) =>
        {
            string text = lblImodMap.Tag as string ?? lblImodMap.Text;
            if (!string.IsNullOrWhiteSpace(text))
            {
                Clipboard.SetText(text);
                ShowCopiedToolTip(lblImodMap);
            }
        };

        const string imodRoleTemplate = "Mouse=0x0, Keyboard=0xC8, Audio=0xFA0, Gamepad=0xC8";
        Label lblImodHelp = new()
        {
            Text = "input: hex / list / roles",
            AutoSize = false,
            AutoEllipsis = true,
            ForeColor = _mutedText,
        };

        IReadOnlyList<string> imodDeviceEditorRoles = GetImodDeviceEditorRoles(device);
        Dictionary<string, ThemedTextBox> imodDeviceEditorBoxes = new(StringComparer.OrdinalIgnoreCase);
        bool suppressImodDeviceEditor = false;
        Panel imodDeviceEditorPanel = new()
        {
            BackColor = _bgGroup,
            Visible = imodDeviceEditorRoles.Count > 0,
        };

        void SyncImodDeviceEditorFromText()
        {
            if (imodDeviceEditorBoxes.Count == 0 || suppressImodDeviceEditor)
            {
                return;
            }

            bool hasRoleValues = TryParseImodRoleIntervals(txtImod.Text ?? string.Empty, out Dictionary<string, uint> roleValues)
                && roleValues.Count > 0;
            uint? singleValue = null;
            string imodText = txtImod.Text ?? string.Empty;
            if (!hasRoleValues
                && !string.IsNullOrWhiteSpace(imodText)
                && TryParseImodInterval(imodText, ImodDefaultInterval, out uint parsedSingle))
            {
                singleValue = parsedSingle;
            }

            suppressImodDeviceEditor = true;
            try
            {
                foreach (KeyValuePair<string, ThemedTextBox> pair in imodDeviceEditorBoxes)
                {
                    uint value = hasRoleValues && roleValues.TryGetValue(pair.Key, out uint roleValue)
                        ? roleValue
                        : singleValue ?? GetImodDeviceEditorDefaultValue(pair.Key);
                    pair.Value.Text = FormatImodValue(value);
                    pair.Value.ForeColor = _fgMain;
                }
            }
            finally
            {
                suppressImodDeviceEditor = false;
            }
        }

        void UpdateImodTextFromDeviceEditor()
        {
            if (suppressImodDeviceEditor || imodDeviceEditorBoxes.Count == 0)
            {
                return;
            }

            Dictionary<string, uint> roleValues = new(StringComparer.OrdinalIgnoreCase);
            bool valid = true;
            foreach (string role in imodDeviceEditorRoles)
            {
                if (!imodDeviceEditorBoxes.TryGetValue(role, out ThemedTextBox? box))
                {
                    continue;
                }

                string text = box.Text.Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                if (!TryParseImodInterval(text, ImodDefaultInterval, out uint value))
                {
                    box.ForeColor = Color.FromArgb(255, 110, 110);
                    valid = false;
                    continue;
                }

                box.ForeColor = _fgMain;
                roleValues[role] = value;
            }

            if (!valid || roleValues.Count == 0)
            {
                return;
            }

            suppressImodDeviceEditor = true;
            try
            {
                txtImod.Text = FormatImodRoleIntervals(roleValues);
                chkImod.Checked = true;
                cmbImodMode.SelectedItem = ImodModeRoles;
            }
            finally
            {
                suppressImodDeviceEditor = false;
            }
        }

        void NormalizeImodDeviceEditorBox(ThemedTextBox box)
        {
            if (TryParseImodInterval(box.Text.Trim(), ImodDefaultInterval, out uint value))
            {
                suppressImodDeviceEditor = true;
                try
                {
                    box.Text = FormatImodValue(value);
                    box.ForeColor = _fgMain;
                }
                finally
                {
                    suppressImodDeviceEditor = false;
                }

                UpdateImodTextFromDeviceEditor();
            }
        }

        if (imodDeviceEditorRoles.Count > 0)
        {
            int editorLabelWidth = UiScale(92);
            int editorBoxWidth = UiScale(78);
            int editorColumnGap = UiScale(18);
            int editorRowHeight = UiScale(28);
            // Keep the editor usable on narrower windows: one or two roles fit on one row,
            // while three or more use a stable two-column grid. This prevents the third
            // value box from being clipped when Russian labels are wider.
            bool twoColumnGrid = imodDeviceEditorRoles.Count >= 3;
            int editorX = 0;
            int editorY = 0;
            for (int i = 0; i < imodDeviceEditorRoles.Count; i++)
            {
                if (twoColumnGrid && i > 0 && i % 2 == 0)
                {
                    editorX = 0;
                    editorY += editorRowHeight;
                }

                string role = imodDeviceEditorRoles[i];
                Label roleLabel = new()
                {
                    Text = role + ":",
                    AutoSize = false,
                    Size = new Size(editorLabelWidth, UiScale(22)),
                    Location = new Point(editorX, editorY + UiScale(3)),
                    ForeColor = _fgMain,
                };

                ThemedTextBox roleBox = new()
                {
                    Size = new Size(editorBoxWidth, UiScale(24)),
                    Location = new Point(roleLabel.Right + UiScale(4), editorY),
                    BackColor = Color.FromArgb(18, 18, 22),
                    ForeColor = _fgMain,
                    Text = FormatImodValue(GetImodDeviceEditorDefaultValue(role)),
                    TextAlign = HorizontalAlignment.Left,
                    Font = _blockFont,
                };
                StyleDarkTextBox(roleBox);
                roleBox.TextChanged += (_, _) => UpdateImodTextFromDeviceEditor();
                roleBox.Inner.Leave += (_, _) => NormalizeImodDeviceEditorBox(roleBox);
                roleBox.Inner.KeyDown += (_, e) =>
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        NormalizeImodDeviceEditorBox(roleBox);
                        e.SuppressKeyPress = true;
                    }
                };

                imodDeviceEditorBoxes[role] = roleBox;
                imodDeviceEditorPanel.Controls.AddRange([roleLabel, roleBox]);
                editorX = roleBox.Right + editorColumnGap;
            }
        }

        txtImod.TextChanged += (_, _) => SyncImodDeviceEditorFromText();
        SyncImodDeviceEditorFromText();

        Button btnImodApply = new()
        {
            Text = "SET",
            Size = UiScale(54, 24),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(btnImodApply);
        btnImodApply.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnImodApply);
        btnImodApply.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnImodApply);
        btnImodApply.Click += (_, _) =>
        {
            WriteLog("UI: SET IMOD button clicked");
            DeviceBlock? targetBlock = createdBlock;
            if (targetBlock is not null && targetBlock.Device.IsTestDevice)
            {
                RefreshTestImodPreview(targetBlock, "test-imod-set");
                return;
            }

            if (TryBlockBusyHardwareWrite("IMOD SET") || TryBlockSandboxHardwareWrite("IMOD SET"))
            {
                return;
            }

            OperationReport report = new();
            if (!CreateDeviceTweakerBackup("pre-imod", showDialog: false))
            {
                report.MarkNoChangesMade();
                report.AddError("Automatic backup", "backup could not be created; changes were not applied");
                ShowOperationResult(
                    report,
                    successMessage: string.Empty,
                    partialMessage: "IMOD apply was cancelled because the automatic backup failed.",
                    operationName: "USB IMOD");
                return;
            }

            report.SetBackupPath(_lastBackupPath);

            ImodApplyOutcome outcome = ApplyImodSettings(out string? note, out string? technicalDetails);
            if (outcome is ImodApplyOutcome.SkippedNoUsb or ImodApplyOutcome.SkippedNoController)
            {
                ShowThemedInfo(string.IsNullOrWhiteSpace(note)
                    ? "No eligible USB IMOD target for SET."
                    : note);
                return;
            }

            WaitForBackgroundUiTasks(RefreshImodCurrentValuesAsync(showReadingStatus: true, reason: "imod-set"));
            AddImodResultToReport(report, outcome, note, technicalDetails);

            if (report.Succeeded && report.Warnings.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(note))
                {
                    ShowThemedInfo(note);
                }

                return;
            }

            ShowOperationResult(
                report,
                successMessage: note ?? "IMOD applied.",
                partialMessage: "USB IMOD completed with warnings or failed steps.",
                operationName: "USB IMOD");
            MaybeOfferVulnerableDriverBlocklistDisable(report);
        };

        Button btnImodDelete = new()
        {
            Text = "DELETE",
            Size = UiScale(76, 24),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(btnImodDelete);
        btnImodDelete.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnImodDelete);
        btnImodDelete.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnImodDelete);
        btnImodDelete.Click += (_, _) =>
        {
            WriteLog("UI: DELETE IMOD button clicked");
            DeviceBlock? targetBlock = createdBlock;
            if (targetBlock is not null && targetBlock.Device.IsTestDevice)
            {
                targetBlock.ImodBox.Text = FormatImodValue(ImodDefaultInterval);
                UpdateImodSelectorsFromText(targetBlock);
                RefreshTestImodPreview(targetBlock, "test-imod-delete");
                return;
            }

            if (TryBlockBusyHardwareWrite("IMOD DELETE") || TryBlockSandboxHardwareWrite("IMOD DELETE"))
            {
                return;
            }

            OperationReport report = new();
            try
            {
                ResetImodIntervalsToDefault("imod-delete", report);
            }
            catch (Exception ex)
            {
                WriteLog($"IMOD.DELETE: failed: {ex.Message}");
                report.AddError("IMOD delete", ex.Message);
            }

            ShowOperationResult(
                report,
                successMessage: "IMOD reset to defaults.\nStartup script removed. Reboot your PC to unload DTIMOD.sys from memory if it was loaded.",
                partialMessage: "IMOD delete finished with errors. Some persistence files may still remain. If the driver was loaded, reboot to unload DTIMOD.sys from memory.",
                operationName: "USB IMOD DELETE");
            if (report.Succeeded)
            {
                WriteLog("IMOD.DELETE: done; reboot recommended to unload DTIMOD.sys from memory");
            }
        };

        Button btnImodCheck = new()
        {
            Text = "CHECK",
            Size = UiScale(76, 24),
            FlatStyle = FlatStyle.Flat,
            Font = _blockFont,
            UseVisualStyleBackColor = false,
            Cursor = Cursors.Hand,
        };
        SetTopButtonBaseStyle(btnImodCheck);
        btnImodCheck.MouseEnter += (_, _) => SetTopButtonHoverStyle(btnImodCheck);
        btnImodCheck.MouseLeave += (_, _) => SetTopButtonBaseStyle(btnImodCheck);
        btnImodCheck.Click += (_, _) =>
        {
            DeviceBlock? targetBlock = createdBlock;
            if (targetBlock is null)
            {
                return;
            }

            CheckImodDriverFromBlock(targetBlock);
        };

        void UpdateImodDetailsVisibility()
        {
            bool showDetails = chkImod.Visible
                && (chkImod.Checked || IsImodAttentionStatus(GetSourceControlText(lblImodCurrent)));
            lblImodCurrent.Visible = showDetails;
            lblImodDefault.Visible = showDetails;
            lblImodMap.Visible = showDetails;
        }

        void UpdateImodModeHint()
        {
            string mode = cmbImodMode.SelectedItem?.ToString() ?? ImodModeSingle;
            lblImodModeHint.Text = mode switch
            {
                ImodModeRoles => "detected USB device roles",
                ImodModeVector => "values by interrupter index",
                _ => "single value for controller",
            };
        }

        bool showImod = ShouldShowImod(device);
        int imodRowTop = -1;

        void LayoutImodRow()
        {
            if (!showImod || imodRowTop < 0)
            {
                return;
            }

            // Manually positioned controls still carry a small WinForms
            // margin. Keep it inside the panel instead of ending exactly on
            // the right edge, where layout diagnostics correctly see it as
            // clipped after a language switch.
            int usableSettingsWidth = Math.Max(UiScale(120), availableSettingsWidth - UiScale(4));

            lblImodMode.Location = new Point(0, imodRowTop + labelOffset);
            cmbImodMode.Location = new Point(valueX, imodRowTop);
            int modeHintLeft = cmbImodMode.Right + UiScale(14);
            lblImodModeHint.Location = new Point(modeHintLeft, imodRowTop + labelOffset);
            lblImodModeHint.Size = new Size(
                Math.Max(UiScale(110), usableSettingsWidth - modeHintLeft - UiScale(8)),
                UiScale(18));

            int valueRowTop = cmbImodMode.Bottom + rowGap;
            Size imodLabelSize = TextRenderer.MeasureText(lblImod.Text, _baseFont);
            int checkY = valueRowTop + labelOffset + Math.Max(0, (imodLabelSize.Height - imodCheckSize) / 2);
            chkImod.Location = new Point(0, checkY);
            lblImod.Location = new Point(imodCheckSize + imodCheckGap, valueRowTop + labelOffset);

            int buttonGap = UiScale(6);
            int inlineGap = UiScale(8);
            int buttonTotalWidth = btnImodApply.Width + btnImodDelete.Width + btnImodCheck.Width + (buttonGap * 2);
            int inlineInputWidth = usableSettingsWidth - valueX - buttonTotalWidth - inlineGap;
            bool useInlineImodButtons = inlineInputWidth >= UiScale(160);
            int imodInputWidth = useInlineImodButtons
                ? Math.Min(UiScale(420), inlineInputWidth)
                : Math.Min(UiScale(420), Math.Max(UiScale(160), usableSettingsWidth - valueX - UiScale(8)));
            txtImod.Width = imodInputWidth;
            txtImod.Location = new Point(valueX, valueRowTop);
            lblImodHelp.Visible = false;
            lblImodHelp.Location = new Point(txtImod.Left, txtImod.Bottom + UiScale(4));
            lblImodHelp.Size = new Size(UiScale(180), UiScale(18));

            int imodButtonTop = useInlineImodButtons ? txtImod.Top : txtImod.Bottom + UiScale(6);
            int imodButtonLeft = useInlineImodButtons ? txtImod.Right + inlineGap : 0;
            btnImodApply.Location = new Point(imodButtonLeft, imodButtonTop);
            btnImodDelete.Location = new Point(btnImodApply.Right + buttonGap, imodButtonTop);
            btnImodCheck.Location = new Point(btnImodDelete.Right + buttonGap, imodButtonTop);

            int statusTop = Math.Max(txtImod.Bottom, btnImodCheck.Bottom) + UiScale(6);
            if (imodDeviceEditorPanel.Visible)
            {
                imodDeviceEditorPanel.Location = new Point(valueX, statusTop + UiScale(6));
                int editorRows = imodDeviceEditorRoles.Count >= 3 ? 2 : 1;
                imodDeviceEditorPanel.Size = new Size(
                    Math.Max(UiScale(80), usableSettingsWidth - valueX - UiScale(8)),
                    UiScale(editorRows * 26));
                statusTop = imodDeviceEditorPanel.Bottom + UiScale(6);
            }

            int imodDetailsX = valueX + readbackInset;
            int imodStatusTop = statusTop + UiScale(4);
            lblImodCurrent.Location = new Point(imodDetailsX, imodStatusTop);
            int defaultStatusWidth = UiScale(UiLanguage.IsRussian ? 152 : 120);
            int statusGap = UiScale(14);
            int currentStatusWidth = Math.Max(
                UiScale(90),
                Math.Min(UiScale(260), usableSettingsWidth - imodDetailsX - defaultStatusWidth - statusGap));
            lblImodCurrent.Size = new Size(currentStatusWidth, UiScale(18));
            lblImodDefault.Location = new Point(lblImodCurrent.Right + statusGap, imodStatusTop);
            lblImodDefault.Size = new Size(defaultStatusWidth, UiScale(18));
            int imodMapRows = imodDeviceEditorRoles.Contains("Gamepad", StringComparer.OrdinalIgnoreCase) ? 12 : 11;
            lblImodMap.Location = new Point(imodDetailsX, lblImodCurrent.Bottom + UiScale(14));
            FitImodMapLabel(lblImodMap, usableSettingsWidth - imodDetailsX, imodMapRows);
        }

        if (showImod)
        {
            lblImodMode.Visible = true;
            cmbImodMode.Visible = true;
            lblImodModeHint.Visible = true;

            UpdateImodModeHint();
            imodRowTop = rowTop;
            LayoutImodRow();
            settingsPanel.Controls.AddRange([lblImodMode, cmbImodMode, lblImodModeHint, chkImod, lblImod, txtImod, btnImodApply, btnImodDelete, btnImodCheck, imodDeviceEditorPanel, lblImodCurrent, lblImodDefault, lblImodMap]);
            _copyToolTip.SetToolTip(txtImod, $"Supported IMOD input:\n0xC8\n0xC8, 0xFA0\n{imodRoleTemplate}");
            _copyToolTip.SetToolTip(cmbImodMode, "XHCI applies one value to all interrupters on this USB host controller. Device mapping assigns detected devices to their interrupter. Interrupter mode applies values by interrupter index.");
            _copyToolTip.SetToolTip(lblImodModeHint, "Shows how the selected IMOD mode interprets the IMOD Value field.");
            _copyToolTip.SetToolTip(btnImodApply, "Apply the current IMOD configuration and read back hardware values.");
            _copyToolTip.SetToolTip(btnImodCheck, "Load DTIMOD.sys (IMOD driver) and re-read current hardware IMOD values.");
            _copyToolTip.SetToolTip(lblImodHelp, "Role mode applies values to the detected interrupter for each USB role.");
            _copyToolTip.SetToolTip(lblImodDefault, "Click to copy the default IMOD interval into the input field.");
            _copyToolTip.SetToolTip(lblImodMap, string.Empty);
            UpdateImodDetailsVisibility();
            rowTop = Math.Max(lblImodMap.Bottom, lblImodCurrent.Bottom) + rowGap;
        }
        else
        {
            chkImod.Visible = false;
            lblImod.Visible = false;
            txtImod.Visible = false;
            lblImodMode.Visible = false;
            cmbImodMode.Visible = false;
            lblImodModeHint.Visible = false;
            lblImodDefault.Visible = false;
            lblImodCurrent.Visible = false;
            lblImodMap.Visible = false;
            btnImodApply.Visible = false;
            btnImodDelete.Visible = false;
            btnImodCheck.Visible = false;
        }

        int settingsContentRight = 0;
        int settingsContentBottom = 0;
        foreach (Control child in settingsPanel.Controls)
        {
            if (!child.Visible)
            {
                continue;
            }

            settingsContentRight = Math.Max(settingsContentRight, child.Right);
            settingsContentBottom = Math.Max(settingsContentBottom, child.Bottom);
        }

        desiredSettingsSideWidth = Math.Max(UiScale(600), settingsContentRight + UiScale(8));
        UpdateResponsivePlacement();

        int settingsMinWidth = Math.Min(UiScale(420), availableSettingsWidth);
        Size settingsSize = new(
            Math.Min(Math.Max(settingsContentRight + UiScale(8), settingsMinWidth), availableSettingsWidth),
            Math.Max(settingsContentBottom + UiScale(8), UiScale(24)));
        settingsPanel.Size = settingsSize;

        // Top-align settings with the CPU list when they fit beside dual-CCD
        // columns. Otherwise stack settings under Mask/IRQ so affinity never
        // depends on a horizontal AutoScroll viewport.
        void PlaceAffinityChrome()
        {
            int maskY = cpuPanel.Bottom + UiScale(12);
            lblMask.Location = new Point(lblMask.Left, maskY);
            lblIrq.Location = new Point(lblIrq.Left, maskY + UiScale(20));

            if (affinitySettingsStacked)
            {
                settingsPanel.Location = new Point(settingsX, lblIrq.Bottom + UiScale(12));
            }
            else
            {
                settingsPanel.Location = new Point(settingsX, cpuPanel.Top);
            }
        }

        PlaceAffinityChrome();

        int infoY = Math.Max(lblIrq.Bottom, settingsPanel.Bottom) + UiScale(12);

        InfoTextBox lblInfo = new()
        {
            Text = "PNP ID: -",
            Location = new Point(UiScale(18), infoY),
            Size = new Size(grp.Width - UiScale(40), UiScale(70)),
            // Keep the technical information block aligned with the larger
            // registry/status text used elsewhere in the device card.
            Font = _technicalFont,
            Cursor = Cursors.Hand,
            BackColor = _bgGroup,
            ForeColor = _mutedText,
            PrefixColor = _statusPrefix,
            ValueColor = _fgMain,
            SeparatorColor = _statusSeparator,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            TabStop = false,
        };
        lblInfo.Click += (_, _) =>
        {
            if (lblInfo.Tag is string txt && !string.IsNullOrWhiteSpace(txt))
            {
                Clipboard.SetText(txt);
                ShowCopiedToolTip(lblInfo);
                WriteLog($"UI: Copied registry path: {txt}");
            }
        };
        _copyToolTip.SetToolTip(lblInfo, "Click to copy full registry path to clipboard");

        void RelayoutDeviceBlockChrome()
        {
            UpdateResponsivePlacement();
            LayoutNicItrRow();
            LayoutImodRow();

            void FitCurrentImodMap()
            {
                if (!showImod || !lblImodMap.Visible)
                {
                    return;
                }

                int maxRows = imodDeviceEditorRoles.Contains("Gamepad", StringComparer.OrdinalIgnoreCase) ? 12 : 11;
                int mapWidth = Math.Max(UiScale(120), availableSettingsWidth - lblImodMap.Left - UiScale(4));
                FitImodMapLabel(lblImodMap, mapWidth, maxRows);
            }

            // IMOD readback replaces the initial one-line placeholder with a
            // multi-line device/interrupter map. Refit before measuring the
            // settings panel so the new lines are not trapped in a one-row box.
            FitCurrentImodMap();

            int MeasureVisibleSettings(out int right, out int bottom)
            {
                right = 0;
                bottom = 0;
                foreach (Control child in settingsPanel.Controls)
                {
                    if (!child.Visible)
                    {
                        continue;
                    }

                    right = Math.Max(right, child.Right);
                    bottom = Math.Max(bottom, child.Bottom);
                }

                return bottom;
            }

            MeasureVisibleSettings(out int visibleSettingsRight, out _);
            desiredSettingsSideWidth = Math.Max(UiScale(600), visibleSettingsRight + UiScale(8));
            UpdateResponsivePlacement();
            LayoutImodRow();
            FitCurrentImodMap();

            MeasureVisibleSettings(out visibleSettingsRight, out int visibleSettingsBottom);

            int currentAvailableSettingsWidth = Math.Max(UiScale(260), grp.Width - settingsX - UiScale(24));
            int currentSettingsMinWidth = Math.Min(UiScale(420), currentAvailableSettingsWidth);
            Size currentSettingsSize = new(
                Math.Min(Math.Max(visibleSettingsRight + UiScale(8), currentSettingsMinWidth), currentAvailableSettingsWidth),
                Math.Max(visibleSettingsBottom + UiScale(8), UiScale(24)));
            settingsPanel.Size = currentSettingsSize;

            PlaceAffinityChrome();

            int currentInfoWidth = Math.Max(UiScale(140), grp.Width - lblInfo.Left - UiScale(24));
            int currentInfoHeight = Math.Max(UiScale(70), GetPreferredTextHeight(lblInfo, currentInfoWidth) + UiScale(2));
            lblInfo.Location = new Point(lblInfo.Left, Math.Max(lblIrq.Bottom, settingsPanel.Bottom) + UiScale(12));
            lblInfo.Size = new Size(currentInfoWidth, currentInfoHeight);

            grp.Height = lblInfo.Bottom + UiScale(20);
        }

        List<Control> chrome =
        [
            headerPanel,
            divider,
            cpuLabel,
            cpuPanel,
            lblMask,
            lblIrq,
            settingsPanel,
            lblInfo,
        ];
        if (affinityQuickPanel is not null)
        {
            chrome.Add(affinityQuickPanel);
        }
        grp.Controls.AddRange(chrome.ToArray());
        WireDevicesMouseWheelForwarding(grp);

        RelayoutDeviceBlockChrome();

        DeviceBlock block = new()
        {
            Device = device,
            Kind = device.Kind,
            Group = grp,
            HeaderPanel = headerPanel,
            Divider = divider,
            CpuTitleLabel = cpuLabel,
            CpuPanel = cpuPanel,
            SettingsPanel = settingsPanel,
            TitleLabel = headerLabel,
            CpuBoxes = cpuBoxes,
            AffinityLabel = lblMask,
            IrqLabel = lblIrq,
            MsiCombo = cmbMsi,
            PowerSavingCheck = chkPowerSaving,
            LimitBox = txtLimit.Inner,
            PrioCombo = cmbPrio,
            PolicyCombo = cmbPolicy,
            PolicyLabel = lblPolicy,
            NdisModeLabel = device.Kind == DeviceKind.NET_NDIS ? lblNdisMode : null,
            NdisModeCombo = device.Kind == DeviceKind.NET_NDIS ? cmbNdisMode : null,
            RssQueueBox = device.Kind == DeviceKind.NET_NDIS ? nudRssQueues : null,
            NicItrBox = showNicItr ? txtNicItr.Inner : null,
            NicItrStatusLabel = showNicItr ? lblNicItrStatus : null,
            NicItrTimeLabel = showNicItr ? lblNicItrTime : null,
            NicItrApplyButton = showNicItr ? btnNicItr : null,
            NicItrSaveButton = showNicItr ? btnNicItrSave : null,
            NicItrCheckButton = showNicItr ? btnNicItrCheck : null,
            ImodAutoCheck = chkImod,
            ImodModeCombo = showImod ? cmbImodMode : null,
            ImodCheckButton = showImod ? btnImodCheck : null,
            ImodBox = txtImod.Inner,
            ImodDefaultLabel = lblImodDefault,
            ImodCurrentLabel = lblImodCurrent,
            ImodMapLabel = lblImodMap,
            RawMouseThrottleCheck = showRawMouseThrottle ? chkRawMouseThrottle : null,
            RawMouseThrottleCombo = showRawMouseThrottle ? cmbRawMouseThrottle : null,
            RawMouseThrottleStatusLabel = showRawMouseThrottle ? lblRawMouseThrottleStatus : null,
            InfoLabel = lblInfo,
            RelayoutAction = RelayoutDeviceBlockChrome,
            AffinityMask = 0,
            IrqCount = null,
            ModifiedBadge = modifiedBadge,
            HeaderNote = headerNote,
        };
        createdBlock = block;

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
                int queues = ClampRssQueueCount((int)block.RssQueueBox.Value);
                ApplyNdisSelection(block, baseCore, queues);
            };
        }

        if (block.NdisModeCombo is not null)
        {
            block.NdisModeCombo.SelectedIndexChanged += (_, _) =>
            {
                RecalcAffinityMask(block);
                UpdateBlockInfoText(block);
            };
        }

        if (block.RawMouseThrottleCheck is not null && block.RawMouseThrottleCombo is not null)
        {
            block.RawMouseThrottleCheck.CheckedChanged += (_, _) => HandleRawMouseThrottleChanged(block);
            block.RawMouseThrottleCombo.SelectedIndexChanged += (_, _) => HandleRawMouseThrottleChanged(block);
        }

        if (block.NicItrApplyButton is not null)
        {
            block.NicItrApplyButton.Click += (_, _) => ApplyNicItrFromBlock(block);
        }

        if (block.NicItrSaveButton is not null)
        {
            block.NicItrSaveButton.Click += (_, _) => SaveNicItrPersistenceFromBlock(block);
        }

        if (block.NicItrCheckButton is not null)
        {
            block.NicItrCheckButton.Click += (_, _) => CheckImodDriverFromNicBlock(block);
        }

        if (block.NicItrBox is not null)
        {
            block.NicItrBox.TextChanged += (_, _) => UpdateNicItrInputTimeLabel(block);
        }

        if (showImod)
        {
            InitializeImodSelectors(block);
            block.ImodAutoCheck.CheckedChanged += (_, _) =>
            {
                UpdateImodDetailsVisibility();
                RelayoutDeviceBlockChrome();
                LayoutBlocks();
            };
            block.ImodModeCombo!.SelectedIndexChanged += (_, _) =>
            {
                HandleImodModeChanged(block);
                UpdateImodModeHint();
            };
            block.ImodBox.TextChanged += (_, _) =>
            {
                if (block.SuppressImodEvents == 0)
                {
                    UpdateImodSelectorsFromText(block);
                }
            };
        }

        LoadBlockSettings(block);
        if (device.Wifi)
        {
            // Wi-Fi drivers are intentionally left completely untouched. Keep
            // detected values visible, but do not expose controls that APPLY
            // and AUTO will never write.
            settingsPanel.Enabled = false;
            cpuPanel.Enabled = false;
            _copyToolTip.SetToolTip(settingsPanel, "Wi-Fi settings are preserved by DEVICE TWEAKER.");
            _copyToolTip.SetToolTip(cpuPanel, "Wi-Fi settings are preserved by DEVICE TWEAKER.");
        }
        if (showImod)
        {
            UpdateImodDetailsVisibility();
            RelayoutDeviceBlockChrome();
        }

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
        if (showImod)
        {
            block.ImodAutoCheck.CheckedChanged += (_, _) => OnBlockSettingChanged(block);
            if (block.ImodBox is not null)
            {
                block.ImodBox.TextChanged += (_, _) => OnBlockSettingChanged(block);
            }
        }

        block.CaptureInitialState();

        _devicesPanel.Controls.Add(grp);
        _blocks.Add(block);
        if (showImod && block.Device.IsTestDevice)
        {
            RefreshTestImodPreview(block, "test-block-init");
        }
        if (showNicItr)
        {
            RefreshNicItrBlock(block);
        }
    }

    private void LayoutBlocks()
    {
        int paddingX = UiScale(24);
        int gapY = UiScale(18);
        int y = UiScale(18);
        bool firstPlaced = true;

        bool showReserved = _reservedCpuPanel is not null
            && string.Equals(_activeCategoryFilter, "ALL", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(_searchFilterText);
        if (_reservedCpuPanel is not null)
        {
            _reservedCpuPanel.Visible = showReserved;
        }
        Panel? reserved = showReserved ? _reservedCpuPanel : null;
        bool reservedInserted = false;

        DeviceBlock? lastStorBlock = null;
        if (reserved is not null && _blocks.Count > 0)
        {
            List<DeviceBlock> storBlocks = _blocks.Where(b => b.Kind == DeviceKind.STOR).ToList();
            if (storBlocks.Count > 0)
            {
                lastStorBlock = storBlocks[^1];
            }
        }

        DeviceBlock? lastBlock = _blocks.Count > 0 ? _blocks[^1] : null;

        void PlaceControl(Control control)
        {
            if (!firstPlaced)
            {
                y += gapY;
            }

            control.Location = new Point(paddingX, y);
            y += control.Height;
            firstPlaced = false;
        }

        foreach (DeviceBlock b in _blocks)
        {
            if (!b.Group.Visible)
            {
                continue;
            }
            int width = GetDevicesViewportWidth() - paddingX - UiScale(12);
            if (width < UiScale(360))
            {
                width = UiScale(360);
            }

            b.Group.Width = width;
            RelayoutDeviceBlockChrome(b);

            int currentHeight = b.InfoLabel.Height > 0 ? b.InfoLabel.Height : UiScale(60);
            int infoWidth = Math.Max(UiScale(140), b.Group.Width - b.InfoLabel.Left - UiScale(24));
            b.InfoLabel.Size = new Size(infoWidth, currentHeight);

            PlaceControl(b.Group);

            if (reserved is not null && !reservedInserted)
            {
                bool placeHere = false;
                if (lastStorBlock is not null && ReferenceEquals(b, lastStorBlock))
                {
                    placeHere = true;
                }
                else if (lastStorBlock is null && lastBlock is not null && ReferenceEquals(b, lastBlock))
                {
                    placeHere = true;
                }

                if (placeHere)
                {
                    UpdateReservedCpuSetsPanelLayout(reserved, width);
                    PlaceControl(reserved);
                    reservedInserted = true;
                }
            }
        }

        if (reserved is not null && !reservedInserted)
        {
            int width = GetDevicesViewportWidth() - paddingX - UiScale(12);
            if (width < UiScale(360))
            {
                width = UiScale(360);
            }

            UpdateReservedCpuSetsPanelLayout(reserved, width);
            PlaceControl(reserved);
        }

        int bottomPadding = UiScale(32);
        int contentHeight = y + bottomPadding;
        if (_devicesPanel.Height != contentHeight)
        {
            _devicesPanel.Height = contentHeight;
        }

        if (_noMatchesLabel is not null)
        {
            _noMatchesLabel.Visible = firstPlaced && _blocks.Count > 0;
            if (_noMatchesLabel.Visible)
            {
                _noMatchesLabel.BringToFront();
            }
        }

        SyncDevicesScrollBar();
    }

    private void EnsureDevicesBusyOverlay()
    {
        if (_devicesBusyOverlay is not null || _devicesHost is null)
        {
            return;
        }

        _devicesBusyOverlay = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = _bgForm,
            Visible = false,
            TabStop = false,
        };

        _devicesBusyLabel = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = _fgMain,
            BackColor = _bgForm,
            Font = _dialogFont,
            Text = string.Empty,
        };

        _devicesBusyOverlay.Controls.Add(_devicesBusyLabel);
        _devicesHost.Controls.Add(_devicesBusyOverlay);
        _devicesBusyOverlay.BringToFront();
    }

    private void BeginDevicesBusy(string stage, int percent)
    {
        if (IsDisposed || !IsHandleCreated)
        {
            return;
        }

        EnsureDevicesBusyOverlay();
        if (_devicesBusyOverlay is null || _devicesBusyLabel is null)
        {
            return;
        }

        _devicesBusyDepth++;
        if (_devicesBusyDepth == 1)
        {
            SetOperationButtonsEnabled(false);
        }
        percent = Math.Clamp(percent, 0, 100);
        _devicesBusyLabel.Text = $"{stage}\r\n{percent}%";
        _devicesBusyOverlay.Visible = true;
        _devicesBusyOverlay.BringToFront();
        _devicesBusyOverlay.Refresh();
        _devicesBusyLabel.Refresh();
    }

    private void BeginDevicesBusyWork(string stage, int totalUnits)
    {
        _devicesBusyDone = 0;
        _devicesBusyTotal = Math.Max(1, totalUnits);
        BeginDevicesBusy(stage, 0);
    }

    private void SetDevicesBusyWork(int totalUnits, int doneUnits)
    {
        _devicesBusyTotal = Math.Max(1, totalUnits);
        _devicesBusyDone = Math.Clamp(doneUnits, 0, _devicesBusyTotal);
    }

    private int GetDevicesBusyPercent()
    {
        if (_devicesBusyTotal <= 0)
        {
            return 0;
        }

        return (int)Math.Clamp(Math.Round(100.0 * _devicesBusyDone / _devicesBusyTotal), 0, 100);
    }

    private void SetDevicesBusyStage(string stage)
    {
        UpdateDevicesBusy(stage, GetDevicesBusyPercent());
    }

    private void TickDevicesBusy(string stage, int units = 1)
    {
        if (units < 0)
        {
            units = 0;
        }

        _devicesBusyDone = Math.Min(_devicesBusyTotal, _devicesBusyDone + units);
        UpdateDevicesBusy(stage, GetDevicesBusyPercent());
    }

    private void UpdateDevicesBusy(string stage, int percent)
    {
        if (_devicesBusyOverlay is null || _devicesBusyLabel is null)
        {
            return;
        }

        percent = Math.Clamp(percent, 0, 100);
        _devicesBusyLabel.Text = $"{stage}\r\n{percent}%";
        if (!_devicesBusyOverlay.Visible)
        {
            return;
        }

        _devicesBusyOverlay.BringToFront();
        _devicesBusyOverlay.Refresh();
        _devicesBusyLabel.Refresh();
    }

    private void EndDevicesBusy()
    {
        if (_devicesBusyOverlay is null)
        {
            return;
        }

        if (_devicesBusyDepth > 0)
        {
            _devicesBusyDepth--;
        }

        if (_devicesBusyDepth > 0)
        {
            return;
        }

        _devicesBusyDepth = 0;
        _devicesBusyDone = 0;
        _devicesBusyTotal = 1;
        _devicesBusyOverlay.Visible = false;
        SetOperationButtonsEnabled(true);
    }

    /// <summary>
    /// Hide progress overlay before modal results so the user never sees
    /// "completed" over a mid-stage percent.
    /// </summary>
    private void CloseDevicesBusyOverlay()
    {
        _devicesBusyDepth = 0;
        _devicesBusyDone = 0;
        _devicesBusyTotal = 1;
        if (_devicesBusyOverlay is not null)
        {
            _devicesBusyOverlay.Visible = false;
        }
        SetOperationButtonsEnabled(true);
    }

    private void SetOperationButtonsEnabled(bool enabled)
    {
        foreach (Button button in _operationButtons)
        {
            if (!button.IsDisposed)
            {
                button.Enabled = enabled;
            }
        }
    }

    private bool TryBlockBusyHardwareWrite(string operation)
    {
        if (_devicesBusyDepth <= 0)
        {
            return false;
        }

        WriteLog($"UI.BUSY.BLOCK: operation=\"{operation}\" depth={_devicesBusyDepth}");
        return true;
    }

    private void WaitForBackgroundUiTasks(params Task[] tasks)
    {
        if (tasks.Length == 0)
        {
            return;
        }

        Task all = Task.WhenAll(tasks);
        if (all.IsCompleted)
        {
            all.GetAwaiter().GetResult();
            return;
        }

        bool wasButtonsEnabled = _operationButtons.Length > 0 && _operationButtons[0].Enabled;
        if (wasButtonsEnabled)
        {
            SetOperationButtonsEnabled(false);
        }
        _devicesBusyDepth++;

        Stopwatch timeout = Stopwatch.StartNew();
        try
        {
            while (!all.IsCompleted)
            {
                if (timeout.Elapsed >= TimeSpan.FromSeconds(20))
                {
                    // Both asynchronous readers validate their generation before touching
                    // controls. Invalidate late results and keep the main window usable if
                    // a WMI/provider call gets stuck inside Windows.
                    Interlocked.Increment(ref _irqRefreshGeneration);
                    Interlocked.Increment(ref _imodReadbackGeneration);
                    WriteLog($"UI.BACKGROUND.TIMEOUT: tasks={tasks.Length} elapsedMs={timeout.ElapsedMilliseconds} late results invalidated");
                    _ = all.ContinueWith(
                        completed =>
                        {
                            if (completed.Exception is not null)
                            {
                                WriteLog($"UI.BACKGROUND.LATE.ERROR: {FlattenLogText(completed.Exception.ToString())}");
                            }
                        },
                        CancellationToken.None,
                        TaskContinuationOptions.ExecuteSynchronously,
                        TaskScheduler.Default);
                    return;
                }

                Application.DoEvents();
                Thread.Sleep(10);
            }

            all.GetAwaiter().GetResult();
        }
        finally
        {
            if (_devicesBusyDepth > 0)
            {
                _devicesBusyDepth--;
            }

            if (wasButtonsEnabled && _devicesBusyDepth == 0)
            {
                SetOperationButtonsEnabled(true);
            }
        }
    }

    private void RefreshBlocks(bool includeImodReadback = true)
    {
        long refreshStarted = Stopwatch.GetTimestamp();
        WriteLog($"REFRESH.START: includeImodReadback={includeImodReadback} previousBlocks={_blocks.Count}");
        bool ownsBusy = _devicesBusyDepth == 0;
        // Temporary budget until device count is known after enumeration.
        if (ownsBusy)
        {
            BeginDevicesBusyWork("Scanning devices...", 8);
        }
        else
        {
            SetDevicesBusyStage("Scanning devices...");
        }

        try
        {
            InvalidateImodCache();
            _ndisRssRuntimeCache.Clear();
            Dictionary<string, string> priorImodStatuses = [];
            foreach (DeviceBlock block in _blocks)
            {
                if (block.ImodCurrentLabel is null || string.IsNullOrWhiteSpace(block.Device.InstanceId))
                {
                    continue;
                }

                string statusText = block.ImodCurrentLabel.Text ?? string.Empty;
                if (statusText.Equals("current: reading...", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                priorImodStatuses[NormalizeInstanceId(block.Device.InstanceId)] = statusText;
            }

            TickDevicesBusy("Clearing device list...", ownsBusy ? 1 : 0);
            _devicesPanel.SuspendLayout();
            try
            {
                // ControlCollection.Clear() only detaches controls. It does not
                // release their native handles, fonts or event subscriptions.
                // A REFRESH therefore has to dispose the complete previous card
                // tree before the new device list is built.
                foreach (Control control in _devicesPanel.Controls.Cast<Control>().ToArray())
                {
                    control.Dispose();
                }

                _devicesPanel.Controls.Clear();
                _devicesPanel.Location = new Point(0, 0);
                if (_devicesScroll is not null)
                {
                    _devicesScroll.Value = 0;
                }
                _blocks.Clear();
                _reservedCpuPanel = null;

                TickDevicesBusy("Enumerating devices...", ownsBusy ? 1 : 0);
                List<DeviceInfo> devs = GetDeviceList();
                WarnIfMissingGpuDriver(devs);

                int buildUnits = Math.Max(0, devs.Count);
                int tailUnits = 2 // reserved CPU + layout
                    + (includeImodReadback ? 1 : 0)
                    + 1; // IRQ
                if (ownsBusy)
                {
                    SetDevicesBusyWork(_devicesBusyDone + buildUnits + tailUnits, _devicesBusyDone);
                }

                int index = 0;
                int total = Math.Max(1, devs.Count);
                if (devs.Count == 0 && ownsBusy)
                {
                    SetDevicesBusyStage("Building device list (0/0)");
                }

                foreach (DeviceInfo d in devs)
                {
                    if (ownsBusy)
                    {
                        TickDevicesBusy($"Building device list ({index + 1}/{total})", 1);
                    }
                    else
                    {
                        SetDevicesBusyStage($"Building device list ({index + 1}/{total})");
                    }

                    NewDeviceBlock(d, index, priorImodStatuses);
                    index++;
                }

                if (ownsBusy)
                {
                    TickDevicesBusy("Building reserved CPU sets...", 1);
                }
                else
                {
                    SetDevicesBusyStage("Building reserved CPU sets...");
                }

                _reservedCpuPanel = NewReservedCpuSetsPanel();
                if (_reservedCpuPanel is not null)
                {
                    _devicesPanel.Controls.Add(_reservedCpuPanel);
                }

                if (ownsBusy)
                {
                    TickDevicesBusy("Laying out devices...", 1);
                }
                else
                {
                    SetDevicesBusyStage("Laying out devices...");
                }

                LayoutBlocks();
            }
            finally
            {
                _devicesPanel.ResumeLayout();
            }

            AdjustInitialDeviceViewportHeight();
            _devicesPanel.Invalidate(true);
            _devicesHost.Invalidate(true);

            if (includeImodReadback)
            {
                // REFRESH no longer loads DTIMOD; this is UI/cache/preview refresh.
                if (ownsBusy)
                {
                    TickDevicesBusy("Updating IMOD display...", 1);
                }
                else
                {
                    SetDevicesBusyStage("Updating IMOD display...");
                }

                WaitForBackgroundUiTasks(RefreshImodCurrentValuesAsync(showReadingStatus: true, reason: "refresh-blocks"));
            }

            if (ownsBusy)
            {
                TickDevicesBusy("Updating IRQ counts...", 1);
            }
            else
            {
                SetDevicesBusyStage("Updating IRQ counts...");
            }

            WaitForBackgroundUiTasks(CalculateIrqCountsAsync("refresh-blocks"));
            LogGuiSnapshot("refresh");
            if (ownsBusy)
            {
                _devicesBusyDone = _devicesBusyTotal;
                UpdateDevicesBusy("Ready", 100);
            }

            UpdateAllBlocksInitialState();
            RebuildFilterCategoryButtons();
            ApplyDeviceFilter();
            UpdateApplyButtonDirtyCount();

            WriteLog(
                $"REFRESH.DONE: includeImodReadback={includeImodReadback} blocks={_blocks.Count} " +
                $"elapsedMs={Stopwatch.GetElapsedTime(refreshStarted).TotalMilliseconds:0}");
        }
        finally
        {
            if (ownsBusy)
            {
                EndDevicesBusy();
            }
        }
    }

    private void WarnIfMissingGpuDriver(IReadOnlyList<DeviceInfo> devices)
    {
        if (_testDevicesEnabled && _testDevicesOnly)
        {
            return;
        }

        bool hasDriver = HasAmdOrNvidiaGpu(devices);
        if (_lastGpuDriverDetected == hasDriver)
        {
            return;
        }

        _lastGpuDriverDetected = hasDriver;
        if (hasDriver)
        {
            _pendingGpuDriverWarning = false;
            return;
        }

        if (!Visible || !IsHandleCreated)
        {
            _pendingGpuDriverWarning = true;
            return;
        }

        ShowMissingGpuDriverWarning();
    }

    private static bool HasAmdOrNvidiaGpu(IReadOnlyList<DeviceInfo> devices)
    {
        foreach (DeviceInfo device in devices)
        {
            if (device.IsTestDevice || device.Kind != DeviceKind.GPU)
            {
                continue;
            }

            string id = device.InstanceId ?? string.Empty;
            if (id.Contains("VEN_10DE", StringComparison.OrdinalIgnoreCase)
                || id.Contains("VEN_1002", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private List<string> BuildIrqLookupKeys(DeviceBlock block)
    {
        List<string> keys = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        void AddKey(string key)
        {
            if (!string.IsNullOrWhiteSpace(key) && seen.Add(key))
            {
                keys.Add(key);
            }
        }

        AddKey(GetIrqPnpKey(block.Device.InstanceId));
        AddKey(GetIrqPnpKey(block.Device.RegBase));
        AddKey(GetIrqPnpKey(GetDisplayRegPath(block.Device.InstanceId)));

        return keys;
    }

    private bool TryFindIrqInfo(
        DeviceBlock block,
        IReadOnlyDictionary<string, DeviceIrqInfo> irqCounts,
        out DeviceIrqInfo? info,
        out string matchedKey)
    {
        foreach (string key in BuildIrqLookupKeys(block))
        {
            if (irqCounts.TryGetValue(key, out info))
            {
                matchedKey = key;
                return true;
            }
        }

        info = null;
        matchedKey = string.Empty;
        return false;
    }

    private static DeviceIrqInfo CreateTestIrqInfo(DeviceBlock block)
    {
        DeviceIrqInfo info = new() { Source = "TEST" };
        int count = block.Device.TestIrqCount ?? block.Kind switch
        {
            DeviceKind.GPU => block.Device.IsIntegratedGpu ? 1 : 2,
            DeviceKind.NET_NDIS or DeviceKind.NET_CX => Math.Max(1, block.RssQueueBox?.Value is decimal queues ? (int)queues : 1),
            DeviceKind.USB => string.IsNullOrWhiteSpace(block.Device.UsbRoles) ? 1 : Math.Min(8, Math.Max(1, block.Device.UsbRoles.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Length)),
            DeviceKind.AUDIO => string.IsNullOrWhiteSpace(block.Device.AudioEndpoints) ? 1 : Math.Min(4, Math.Max(1, block.Device.AudioEndpoints.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Length)),
            DeviceKind.STOR => 1,
            _ => 1,
        };
        count = Math.Clamp(count, 0, 64);
        bool msiEnabled = !block.Device.TestMsiStatus.Equals("Disabled", StringComparison.OrdinalIgnoreCase);

        for (int i = 0; i < count; i++)
        {
            info.AddIrq(msiEnabled ? 1000 + i : i);
        }

        if (count == 0)
        {
            // Auto / Enabled both mean MSI is on in test presets; only Disabled stays Disabled.
            info.MsiStatus = block.Device.TestMsiStatus.Equals("Disabled", StringComparison.OrdinalIgnoreCase)
                ? "Disabled"
                : "Enabled";
        }
        else if (!block.Device.TestMsiStatus.Equals("Auto", StringComparison.OrdinalIgnoreCase))
        {
            info.MsiStatus = block.Device.TestMsiStatus;
        }

        return info;
    }

    private static bool IsKnownMsiStatus(string status)
    {
        return status.Equals("Enabled", StringComparison.OrdinalIgnoreCase)
            || status.Equals("Disabled", StringComparison.OrdinalIgnoreCase);
    }

    private async Task CalculateIrqCountsAsync(string reason = "refresh")
    {
        int generation = ++_irqRefreshGeneration;
        foreach (DeviceBlock b in _blocks)
        {
            b.IrqLabel.Text = "IRQ Count: reading...";
        }

        WriteLog($"IRQ.REFRESH: reason={reason} blocks={_blocks.Count}");
        Dictionary<string, DeviceIrqInfo> irqCounts;
        try
        {
            irqCounts = _blocks.All(block => block.Device.IsTestDevice)
                ? new Dictionary<string, DeviceIrqInfo>(StringComparer.OrdinalIgnoreCase)
                : await Task.Run(GetDeviceIrqCounts);
        }
        catch (Exception ex)
        {
            WriteLog($"IRQ.MAP.WARN: failed to read active IRQ data; using registry fallback: {ex.Message}");
            if (IsDisposed || generation != _irqRefreshGeneration)
            {
                return;
            }

            foreach (DeviceBlock b in _blocks)
            {
                string regMsiStatus = ReadMsiStatusFromRegistry(b);
                b.IrqCount = null;
                b.IrqLabel.Text = $"IRQ Count: unavailable (MSI: {regMsiStatus})";
            }
            return;
        }

        if (IsDisposed || generation != _irqRefreshGeneration)
        {
            return;
        }

        foreach (DeviceBlock b in _blocks)
        {
            string shortPnp = GetShortPnpId(b.Device.InstanceId);
            if (b.Device.IsTestDevice)
            {
                DeviceIrqInfo info = CreateTestIrqInfo(b);
                b.IrqCount = info.Count;
                b.IrqLabel.Text = $"IRQ Count: {info.Count} (MSI: {info.MsiStatus})";
                WriteLog($"IRQ.MAP.TEST: {b.Device.InstanceId} ({shortPnp}) -> count={info.Count} activeMsi={info.MsiStatus} irqs=[{FormatIrqNumbers(info.IrqNumbers)}] source={info.Source}");
            }
            else if (TryFindIrqInfo(b, irqCounts, out DeviceIrqInfo? info, out string matchedKey) && info is not null)
            {
                string regMsiStatus = ReadMsiStatusFromRegistry(b);
                b.IrqCount = info.Count;
                b.IrqLabel.Text = $"IRQ Count: {info.Count} (MSI: {info.MsiStatus})";
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> count={info.Count} activeMsi={info.MsiStatus} registryMsi={regMsiStatus} irqs=[{FormatIrqNumbers(info.IrqNumbers)}] key={matchedKey} source={info.Source}");
                if (IsKnownMsiStatus(info.MsiStatus)
                    && IsKnownMsiStatus(regMsiStatus)
                    && !info.MsiStatus.Equals(regMsiStatus, StringComparison.OrdinalIgnoreCase))
                {
                    WriteLog($"IRQ.MSI.MISMATCH: {b.Device.InstanceId} ({shortPnp}) activeMsi={info.MsiStatus} registryMsi={regMsiStatus} activeIrqs=[{FormatIrqNumbers(info.IrqNumbers)}] key={matchedKey}");
                }
            }
            else
            {
                string regMsiStatus = ReadMsiStatusFromRegistry(b);
                bool isNetwork = b.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX;
                b.IrqCount = isNetwork ? null : 0;
                b.IrqLabel.Text = isNetwork
                    ? $"IRQ Count: N/A (MSI: {regMsiStatus})"
                    : $"IRQ Count: 0 (MSI: {regMsiStatus})";
                string triedKeys = string.Join(", ", BuildIrqLookupKeys(b));
                WriteLog($"IRQ.MAP: {b.Device.InstanceId} ({shortPnp}) -> {(isNetwork ? "N/A" : "0")} activeMsi=Unknown registryMsi={regMsiStatus} irqs=[none] source=registry-fallback keys=[{triedKeys}]");
            }
        }

        LogGuiSnapshot("irq-refresh");
    }

    private void CalculateIrqCounts(string reason = "refresh")
    {
        WaitForBackgroundUiTasks(CalculateIrqCountsAsync(reason));
    }
}
