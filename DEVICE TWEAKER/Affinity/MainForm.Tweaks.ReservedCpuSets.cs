using Microsoft.Win32;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool[] GetReservedCpuSets(int count)
    {
        if (count <= 0)
        {
            return [];
        }

        bool[] reserved = new bool[count];
        string keyPath = @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel";
        const string valueName = "ReservedCpuSets";
        string rawHex = string.Empty;
        List<int> rawIds = [];

        try
        {
            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(keyPath);
            if (key?.GetValue(valueName) is byte[] bytes)
            {
                rawHex = string.Join(" ", bytes.Select(b => b.ToString("X2")));
                int bitIndex = 0;
                foreach (byte b in bytes)
                {
                    for (int i = 0; i < 8; i++)
                    {
                        if ((b & (1 << i)) != 0)
                        {
                            rawIds.Add(bitIndex);
                        }

                        bitIndex++;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            WriteLog($"RESERVED.ERROR: failed to read ReservedCpuSets: {ex.Message}");
        }

        List<int> setBits = [];
        List<int> mappedIds = [];
        List<int> unmappedIds = [];
        foreach (int id in rawIds.Distinct().OrderBy(x => x))
        {
            if (id >= 0 && id < reserved.Length)
            {
                reserved[id] = true;
                setBits.Add(id);
            }
            else if (_cpuIndexByCpuSetId.TryGetValue(id, out int index))
            {
                if (index >= 0 && index < reserved.Length)
                {
                    reserved[index] = true;
                    setBits.Add(index);
                    mappedIds.Add(id);
                }
                else
                {
                    unmappedIds.Add(id);
                }
            }
            else
            {
                unmappedIds.Add(id);
            }
        }

        for (int i = 0; i < reserved.Length; i++)
        {
            if (reserved[i])
            {
                if (!setBits.Contains(i))
                {
                    setBits.Add(i);
                }
            }
        }

        WriteLog(
            $"RESERVED.READ: path=HKLM\\{keyPath}\\{valueName} count={reserved.Length} ids=[{string.Join(',', rawIds.Distinct().OrderBy(x => x))}] mapped=[{string.Join(',', setBits.OrderBy(x => x))}] unmapped=[{string.Join(',', unmappedIds.OrderBy(x => x))}] bytes=[{rawHex}]");
        return reserved;
    }

    private void SetReservedCpuSets(bool[] bits)
    {
        if (bits.Length == 0)
        {
            return;
        }

        string keyPath = @"SYSTEM\CurrentControlSet\Control\Session Manager\Kernel";
        const string valueName = "ReservedCpuSets";
        bool hasAny = bits.Any(b => b);
        List<int> setBits = [];
        for (int i = 0; i < bits.Length; i++)
        {
            if (!bits[i])
            {
                continue;
            }

            setBits.Add(i);
        }

        int maxIndex = setBits.Count > 0 ? setBits.Max() : -1;
        int byteCount = maxIndex >= 0 ? (maxIndex / 8) + 1 : 0;
        byte[] bytes = new byte[byteCount];

        foreach (int id in setBits)
        {
            if (id < 0)
            {
                continue;
            }

            int byteIndex = id / 8;
            int bitIndex = id % 8;
            if (byteIndex >= 0 && byteIndex < bytes.Length)
            {
                bytes[byteIndex] = (byte)(bytes[byteIndex] | (1 << bitIndex));
            }
        }

        try
        {
            using RegistryKey key = Registry.LocalMachine.CreateSubKey(keyPath) ?? throw new InvalidOperationException("Failed to open HKLM key");
            if (!hasAny)
            {
                key.DeleteValue(valueName, throwOnMissingValue: false);
                WriteLog($"RESERVED.WRITE: path=HKLM\\{keyPath}\\{valueName} cleared (no bits set)");
            }
            else
            {
                key.SetValue(valueName, bytes, RegistryValueKind.Binary);
                string hexStr = string.Join(" ", bytes.Select(b => b.ToString("X2")));
                WriteLog($"RESERVED.WRITE: path=HKLM\\{keyPath}\\{valueName} set=[{string.Join(',', setBits)}] bytes=[{hexStr}]");
            }
        }
        catch (Exception ex)
        {
            WriteLog($"RESERVED.ERROR: failed to write ReservedCpuSets: {ex.Message}");
        }
    }

    private void ResetReservedCpuSets()
    {
        if (_reservedCpuPanel?.Tag is not ReservedCpuPanelTag tag || tag.Meta.Count == 0)
        {
            return;
        }

        bool[] empty = new bool[tag.Meta.Count];
        SetReservedCpuSets(empty);

        _suppressReservedCpuEvents++;
        try
        {
            foreach (ReservedCpuEntry entry in tag.Meta)
            {
                entry.Control.Checked = false;
            }
        }
        finally
        {
            _suppressReservedCpuEvents--;
        }

        WriteLog("RESERVED.RESET: cleared ReservedCpuSets via Reset-AllTweaks");
    }

    private void ResetReservedCpuSetsPreview()
    {
        if (_reservedCpuPanel?.Tag is not ReservedCpuPanelTag tag || tag.Meta.Count == 0)
        {
            return;
        }

        _suppressReservedCpuEvents++;
        try
        {
            foreach (ReservedCpuEntry entry in tag.Meta)
            {
                entry.Control.Checked = false;
            }
        }
        finally
        {
            _suppressReservedCpuEvents--;
        }

        WriteLog("RESERVED.DRYRUN: cleared ReservedCpuSets (UI only)");
    }

    private void UpdateReservedCpuSetsRegistry(Panel grp)
    {
        if (grp.Tag is not ReservedCpuPanelTag tag || tag.Meta.Count == 0)
        {
            return;
        }

        if (_suppressReservedCpuEvents > 0)
        {
            return;
        }

        int maxIndex = tag.Meta.Max(e => e.Index);
        if (maxIndex < 0)
        {
            return;
        }

        bool[] bits = new bool[maxIndex + 1];
        foreach (ReservedCpuEntry entry in tag.Meta)
        {
            int iVal = entry.Index;
            if (iVal >= 0 && iVal < bits.Length)
            {
                bits[iVal] = entry.Control.Checked;
            }
        }

        List<int> setBits = [];
        for (int i = 0; i < bits.Length; i++)
        {
            if (bits[i])
            {
                setBits.Add(i);
            }
        }

        WriteLog($"RESERVED.UPDATE: requested set=[{string.Join(',', setBits)}] count={bits.Length}");
        SetReservedCpuSets(bits);
    }

    private Panel? NewReservedCpuSetsPanel()
    {
        int logicalCount = _maxLogical;
        if (_cpuInfo is not null)
        {
            logicalCount = _cpuInfo.Topology.Logical;
        }

        if (logicalCount <= 0)
        {
            logicalCount = Environment.ProcessorCount;
        }

        bool[] reservedBits = GetReservedCpuSets(logicalCount);

        Panel grp = new()
        {
            BackColor = _bgGroup,
            ForeColor = _fgMain,
            Margin = new Padding(0),
            Padding = new Padding(12, 16, 12, 16),
            TabStop = false,
        };
        grp.Paint += (_, e) =>
        {
            Rectangle rect = grp.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };

        Label title = new()
        {
            Text = "Reserved CPU Sets",
            AutoSize = false,
            Font = _titleFont,
            ForeColor = _fgMain,
        };

        Label desc = new()
        {
            Text = @"Reads HKLM:\System\CurrentControlSet\Control\Session Manager\Kernel\ReservedCpuSets",
            AutoSize = false,
            Font = _baseFont,
            ForeColor = _mutedText,
        };

        Label pathLabel = new()
        {
            Text = @"Registry: HKLM\System\CurrentControlSet\Control\Session Manager\kernel",
            Tag = @"HKLM\System\CurrentControlSet\Control\Session Manager\kernel",
            AutoSize = true,
            AutoEllipsis = true,
            Font = _baseFont,
            ForeColor = _fgMain,
            Margin = new Padding(0, 6, 0, 0),
            Cursor = Cursors.Hand,
        };
        pathLabel.MouseEnter += (_, _) => pathLabel.ForeColor = _accent;
        pathLabel.MouseLeave += (_, _) => pathLabel.ForeColor = _fgMain;
        pathLabel.Click += (_, _) =>
        {
            if (pathLabel.Tag is string txt && !string.IsNullOrWhiteSpace(txt))
            {
                Clipboard.SetText(txt);
                ShowCopiedToolTip(pathLabel);
            }
        };

        Panel inner = new()
        {
            BackColor = _bgForm,
            ForeColor = _fgMain,
            AutoScroll = false,
            Margin = new Padding(0, 8, 0, 0),
            Padding = new Padding(8, 6, 8, 6),
        };
        inner.Paint += (_, e) =>
        {
            Rectangle rect = inner.ClientRectangle;
            rect.Width -= 1;
            rect.Height -= 1;
            using Pen pen = new(_border);
            e.Graphics.DrawRectangle(pen, rect);
        };

        List<ReservedCpuEntry> meta = [];
        for (int i = 0; i < logicalCount; i++)
        {
            int eff = 0;
            int ccdId = 0;
            if (_cpuInfo is not null)
            {
                if (_cpuLpByIndex.TryGetValue(i, out CpuLpInfo? lpInfo))
                {
                    eff = lpInfo.EffClass;
                }

                if (_cpuInfo.CcdMap.TryGetValue(i, out int cid))
                {
                    ccdId = cid;
                }
            }

            CheckBox cb = new()
            {
                Text = $"CPU {i}",
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                BackColor = _bgForm,
            };
            StyleCpuCheckbox(cb, i);
            if (reservedBits.Length > i && reservedBits[i])
            {
                cb.Checked = true;
                cb.ForeColor = _accent;
            }

            cb.Tag = grp;
            cb.CheckedChanged += (_, _) =>
            {
                if (cb.Tag is Panel p)
                {
                    UpdateReservedCpuSetsRegistry(p);
                }
            };

            inner.Controls.Add(cb);
            meta.Add(new ReservedCpuEntry { Control = cb, Ccd = ccdId, Eff = eff, Index = i });
        }

        grp.Controls.Add(title);
        grp.Controls.Add(desc);
        grp.Controls.Add(inner);
        grp.Controls.Add(pathLabel);
        grp.Tag = new ReservedCpuPanelTag
        {
            InnerPanel = inner,
            Title = title,
            Description = desc,
            Meta = meta,
            PathLabel = pathLabel,
        };

        return grp;
    }

    private void UpdateReservedCpuSetsPanelLayout(Panel panel, int width)
    {
        if (panel.Tag is not ReservedCpuPanelTag data)
        {
            return;
        }

        panel.Width = width;

        Label title = data.Title;
        Label desc = data.Description;
        Panel inner = data.InnerPanel;
        List<ReservedCpuEntry> meta = data.Meta;
        Label path = data.PathLabel;

        int availWidth = panel.Width - panel.Padding.Left - panel.Padding.Right;
        int y = panel.Padding.Top;

        title.Width = availWidth;
        title.Location = new Point(panel.Padding.Left + 4, y);
        y = title.Bottom + 6;

        desc.Width = availWidth;
        desc.MaximumSize = new Size(availWidth, 0);
        desc.Location = new Point(panel.Padding.Left + 4, y);
        y = desc.Bottom + 8;

        if (meta.Count > 0)
        {
            int startX = 10;
            int columnGap = 16;
            int checkSpacing = 22;
            int runningX = startX;

            List<int> ccdKeys = meta.Select(m => m.Ccd).Distinct().OrderBy(x => x).ToList();
            if (ccdKeys.Count == 0)
            {
                ccdKeys.Add(0);
            }

            List<int> colWidths = [];
            foreach (int cid in ccdKeys)
            {
                List<ReservedCpuEntry> items = meta.Where(m => m.Ccd == cid).ToList();
                List<ReservedCpuEntry> pItems = items.Where(m => m.Eff == 0).ToList();
                List<ReservedCpuEntry> eItems = items.Where(m => m.Eff > 0).ToList();
                List<ReservedCpuEntry> other = items.Except(pItems).Except(eItems).ToList();
                List<ReservedCpuEntry> ordered = [.. pItems, .. eItems, .. other];

                int maxWidth = 120;
                if (ordered.Count > 0)
                {
                    int w = ordered.Max(o => o.Control.PreferredSize.Width);
                    if (w > 0)
                    {
                        maxWidth = w + 10;
                    }
                }

                colWidths.Add(maxWidth);
                int yPos = 4;
                foreach (ReservedCpuEntry entry in ordered)
                {
                    entry.Control.Location = new Point(runningX, yPos);
                    yPos += checkSpacing;
                }

                runningX += maxWidth + columnGap;
            }

            int totalWidthNeeded = startX + colWidths.Sum() + (columnGap * Math.Max(0, colWidths.Count - 1)) + startX;
            inner.Width = Math.Min(availWidth, totalWidthNeeded);

            int maxBottom = meta.Max(m => m.Control.Bottom);
            inner.Height = Math.Max(checkSpacing + 10, maxBottom + 10);
            inner.Location = new Point(panel.Padding.Left + 2, y);
            y = inner.Bottom + 12;
        }

        path.MaximumSize = new Size(availWidth, 0);
        path.Location = new Point(panel.Padding.Left + 4, y);
        y = path.Bottom + panel.Padding.Bottom;

        panel.Height = Math.Max(y, panel.Padding.Vertical + 40);
    }
}
