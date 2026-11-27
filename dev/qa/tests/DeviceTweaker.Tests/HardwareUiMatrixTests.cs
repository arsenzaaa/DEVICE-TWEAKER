using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class HardwareUiMatrixTests
{
    private sealed record CpuProfileSpec(
        string Name,
        CpuInfo Cpu,
        int MaxLogical,
        bool HasHybridCores,
        bool HasDualCcd);

    private sealed record HardwareProfileSpec(
        string Id,
        CpuProfileSpec Cpu,
        DeviceInfo[] Devices);

    [Fact]
    public void AllHardwareProfilesPassLayoutAndInteractiveTransitionsInEnAndRu()
    {
        RunSta(() =>
        {
            UiLanguageCode originalLanguage = UiLanguage.Current;
            try
            {
                HardwareProfileSpec[] profiles = BuildHardwareProfiles();
                UiLanguageCode[] languages = [UiLanguageCode.English, UiLanguageCode.Russian];
                Size[] windowSizes = [new(1172, 875), new(1320, 940)];

                List<string> failures = [];

                foreach (HardwareProfileSpec profile in profiles)
                {
                    foreach (Size windowSize in windowSizes)
                    {
                        UiLanguage.Set(UiLanguageCode.English, persist: false);
                        using MainForm form = new();
                        ConfigureForm(form, profile.Cpu, profile.Devices, windowSize, scale: 1f);

                        foreach (UiLanguageCode language in languages)
                        {
                            UiLanguage.Set(language, persist: false);
                            InvokeLayoutBlocks(form);
                            VerifyAllBlocks(form, $"{profile.Id}/{windowSize.Width}x{windowSize.Height}/{language}/initial", failures);

                            List<DeviceBlock> blocks = GetBlocks(form);
                            foreach (DeviceBlock block in blocks)
                            {
                                VerifyCardInteractions(form, block, $"{profile.Id}/{windowSize.Width}x{windowSize.Height}/{language}", failures);
                            }

                            VerifyAllBlocks(form, $"{profile.Id}/{windowSize.Width}x{windowSize.Height}/{language}/post-interactions", failures);
                        }
                    }
                }

                Assert.True(failures.Count == 0, "Hardware UI matrix failures:\n" + string.Join("\n", failures.Take(30)));
            }
            finally
            {
                UiLanguage.Set(originalLanguage, persist: false);
            }
        });
    }

    [Fact]
    public void ScaleRebindMatrixPreservesLayoutInvariantsAcross100To150Percent()
    {
        RunSta(() =>
        {
            UiLanguageCode originalLanguage = UiLanguage.Current;
            try
            {
                HardwareProfileSpec[] profiles =
                [
                    Build8Core16ThreadGamingProfile(),
                    BuildHybrid14900KProfile(),
                    BuildDualCcd9950X3DProfile(),
                ];

                float[] scales = [1f, 1.25f, 1.5f];
                UiLanguageCode[] languages = [UiLanguageCode.English, UiLanguageCode.Russian];
                List<string> failures = [];

                foreach (HardwareProfileSpec profile in profiles)
                {
                    foreach (UiLanguageCode language in languages)
                    {
                        foreach (float scale in scales)
                        {
                            UiLanguage.Set(language, persist: false);
                            using MainForm form = new();
                            Size scaledWindow = new((int)Math.Round(1172 * scale), (int)Math.Round(875 * scale));
                            ConfigureForm(form, profile.Cpu, profile.Devices, scaledWindow, scale);

                            VerifyAllBlocks(form, $"{profile.Id}/{language}/scale={scale:0.00}", failures);

                            foreach (DeviceBlock block in GetBlocks(form))
                            {
                                VerifyFeedbackStability(form, block, $"{profile.Id}/{language}/scale={scale:0.00}", failures);
                            }

                            // Also exercise 1024x768 viewport clamping (e.g. small displays / CI VMs) at elevated DPI.
                            if (scale > 1f)
                            {
                                form.Size = new Size(1024, 728);
                                InvokeLayoutBlocks(form);
                                VerifyAllBlocks(form, $"{profile.Id}/{language}/scale={scale:0.00}/1024x728", failures);
                                foreach (DeviceBlock block in GetBlocks(form))
                                {
                                    VerifyFeedbackStability(form, block, $"{profile.Id}/{language}/scale={scale:0.00}/1024x728", failures);
                                }
                            }
                        }
                    }
                }

                Assert.True(failures.Count == 0, "DPI scale matrix failures:\n" + string.Join("\n", failures.Take(30)));
            }
            finally
            {
                UiLanguage.Set(originalLanguage, persist: false);
            }
        });
    }

    [Fact]
    public void AutoAffinityPlannerProducesValidAllocationsAcrossAllCpuTopologies()
    {
        CpuProfileSpec[] cpus =
        [
            BuildLaptop4C8T(),
            BuildMidrange6C12T(),
            BuildGaming8C16T(),
            BuildHybrid14600K(),
            BuildHybrid14900K(),
            BuildDualCcd9950X3D(),
            BuildCoreUltra285KNoSmt(),
            BuildServer64Thread(),
        ];

        foreach (CpuProfileSpec cpu in cpus)
        {
            List<AutoCpuUnit> allUnits = BuildAutoCpuUnits(cpu);
            List<AutoCpuUnit> preferred = cpu.HasHybridCores
                ? allUnits.Where(u => !u.IsECore).ToList()
                : (cpu.HasDualCcd ? allUnits.Where(u => u.Ccd == 1).ToList() : allUnits.ToList());
            List<AutoCpuUnit> background = allUnits.Except(preferred).ToList();
            List<int> targetCcdLps = preferred.SelectMany(u => u.Lps).Distinct().OrderBy(x => x).ToList();

            List<AutoAffinityPlanSlot> slots =
            [
                new() { Kind = DeviceKind.GPU, InstanceId = "GPU0", Role = AutoAffinityRole.Gpu, Need = 1, OriginalIndex = 0 },
                new() { Kind = DeviceKind.USB, InstanceId = "USB_MOUSE", Role = AutoAffinityRole.InputMouse, Need = 1, OriginalIndex = 1 },
                new() { Kind = DeviceKind.USB, InstanceId = "USB_KB", Role = AutoAffinityRole.Keyboard, Need = 1, OriginalIndex = 2 },
                new() { Kind = DeviceKind.NET_NDIS, InstanceId = "NIC0", Role = AutoAffinityRole.Nic, Need = 1, OriginalIndex = 3 },
                new() { Kind = DeviceKind.AUDIO, InstanceId = "AUDIO0", Role = AutoAffinityRole.Audio, Need = 1, OriginalIndex = 4 },
            ];

            AutoAffinityPlanner.Allocate(
                slots,
                preferred,
                background,
                allUnits,
                targetCcdLps,
                cpu.MaxLogical,
                visibleCcxSplit: false,
                units => string.Join(",", units.Select(u => u.Id)),
                _ => { });

            foreach (AutoAffinityPlanSlot slot in slots)
            {
                bool mustAssign = slot.Role != AutoAffinityRole.Audio || allUnits.Count >= 5;
                if (mustAssign)
                {
                    Assert.True(slot.Lps.Count >= 1, $"{cpu.Name}: slot {slot.InstanceId} got no LPs");
                }
                Assert.Equal(slot.Lps.Count, slot.Lps.Distinct().Count());
                Assert.All(slot.Lps, lp => Assert.InRange(lp, 0, cpu.MaxLogical - 1));
            }

            if (cpu.HasHybridCores)
            {
                HashSet<int> pCoreLps = preferred.SelectMany(u => u.Lps).ToHashSet();
                foreach (AutoAffinityPlanSlot critical in slots.Where(s => s.Role is AutoAffinityRole.Gpu or AutoAffinityRole.InputMouse or AutoAffinityRole.Nic))
                {
                    Assert.All(critical.Lps, lp => Assert.Contains(lp, pCoreLps));
                }
            }

            int nonCore0Preferred = preferred.Count(u => !u.HasCore0);
            if (nonCore0Preferred >= slots.Count)
            {
                foreach (AutoAffinityPlanSlot critical in slots.Where(s => s.Role is AutoAffinityRole.Gpu or AutoAffinityRole.InputMouse or AutoAffinityRole.Nic))
                {
                    Assert.DoesNotContain(0, critical.Lps);
                }
            }
        }
    }

    [Fact]
    public void DirtyRegistryMaskCorpusDecodesSafelyWithoutThrowing()
    {
        (object? Input, ulong? ExpectedMask)[] corpus =
        [
            (null, null),
            (Array.Empty<byte>(), null),
            (new byte[] { 0x04 }, 0x04UL),
            (new byte[] { 0xFF }, 0xFFUL),
            (new byte[] { 0x00, 0x10 }, 0x1000UL),
            (new byte[] { 0x01, 0x02, 0x04 }, 0x040201UL),
            (new byte[] { 0xAA, 0x00, 0x00, 0x80 }, 0x800000AAUL),
            (new byte[] { 0x01, 0x00, 0x00, 0x00, 0x02 }, 0x0200000001UL),
            (new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40 }, 0x40000000000000UL),
            (new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80 }, 0x8000000000000000UL),
            (new byte[] { 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF }, 0x05UL),
            (4, 4UL),
            (-1, 0xFFFFFFFFUL),
            (0x100000001L, 0x100000001UL),
            (0x8000000000000000UL, 0x8000000000000000UL),
            ("0x10", null),
            (string.Empty, null),
        ];

        foreach (var (input, expectedMask) in corpus)
        {
            ulong? actual = MainForm.TryReadAffinityMaskValue(input);
            Assert.Equal(expectedMask, actual);
        }
    }

    [Fact]
    public void LiveCardMutationsNdisModeAndPopulatedImodMapNeverClip()
    {
        RunSta(() =>
        {
            UiLanguageCode originalLanguage = UiLanguage.Current;
            try
            {
                List<string> failures = [];
                Size[] windowSizes = [new(1000, 720), new(1172, 875)];
                UiLanguageCode[] languages = [UiLanguageCode.English, UiLanguageCode.Russian];

                foreach (Size windowSize in windowSizes)
                {
                    foreach (UiLanguageCode language in languages)
                    {
                        UiLanguage.Set(language, persist: false);
                        using MainForm form = new();
                        HardwareProfileSpec profile = Build8Core16ThreadGamingProfile();
                        ConfigureForm(form, profile.Cpu, profile.Devices, windowSize, scale: 1f);

                        List<DeviceBlock> blocks = GetBlocks(form);
                        DeviceBlock usb = blocks.First(b => b.Kind == DeviceKind.USB);
                        DeviceBlock nic = blocks.First(b => b.Kind == DeviceKind.NET_NDIS && !b.Device.Wifi);

                        // Populate a realistic 12-line IMOD interrupter map and expand DETAILS.
                        if (usb.ImodMapLabel is not null)
                        {
                            usb.ImodReadbackRequested = true;
                            usb.ImodCurrentLabel.Text = "current: varies by interrupter";
                            usb.ImodMapLabel.Text = string.Join(Environment.NewLine,
                            [
                                "devices: Mouse 8K    -> intr0  0x0  0 ns",
                                "         Keyboard 8K -> intr0  0x0  0 ns",
                                "         Microphone  -> intr0  0x0  0 ns",
                                "         Audio       -> intr0  0x0  0 ns",
                                string.Empty,
                                "interrupters 0-1: intr0  0x0   0 ns  | intr1  0xC8  50 us",
                                "interrupters 2-3: intr2  0xC8  50 us | intr3  0xC8  50 us",
                                "interrupters 4-5: intr4  0xC8  50 us | intr5  0xC8  50 us",
                                "interrupters 6-7: intr6  0xC8  50 us | intr7  0xC8  50 us",
                            ]);
                            usb.RestoreImodDetailsExpansion?.Invoke(true);
                            usb.RelayoutAction?.Invoke();
                            VerifySingleBlock(form, usb, $"stress/{windowSize.Width}x{windowSize.Height}/{language}/imod-map-open", failures);
                        }

                        // Switch NDIS Mode live (IRQ -> RSS -> BOTH) which calls UpdateBlockInfoText.
                        if (nic.NdisModeCombo is not null)
                        {
                            for (int modeIdx = 0; modeIdx < nic.NdisModeCombo.Items.Count; modeIdx++)
                            {
                                nic.NdisModeCombo.SelectedIndex = modeIdx;
                                VerifySingleBlock(form, nic, $"stress/{windowSize.Width}x{windowSize.Height}/{language}/ndis-mode-{modeIdx}", failures);
                            }
                        }

                        // Simulate live UpdateBlockInfoText growing InfoLabel after initial build.
                        MethodInfo updateInfo = typeof(MainForm).GetMethod("UpdateBlockInfoText", BindingFlags.Instance | BindingFlags.NonPublic)!;
                        updateInfo.Invoke(form, [usb, "Mouse 8K, Keyboard 8K, Microphone, Audio, Gamepad", "Mouse: 8K, Keyboard: 8K", "Mouse: 8K (confirmed)"]);
                        VerifySingleBlock(form, usb, $"stress/{windowSize.Width}x{windowSize.Height}/{language}/usb-live-info-grow", failures);
                    }
                }

                Assert.True(failures.Count == 0, "Live card mutation failures:\n" + string.Join("\n", failures.Take(30)));
            }
            finally
            {
                UiLanguage.Set(originalLanguage, persist: false);
            }
        });
    }

    [Fact]
    public void DirtyImodInputCorpusNeverThrows()
    {
        MethodInfo tryParse = typeof(MainForm).GetMethod(
            "TryParseImodInput",
            BindingFlags.Static | BindingFlags.NonPublic)!;

        string[] inputs =
        [
            "",
            "   ",
            "0x",
            "0xZZZZ",
            "-1",
            "0x10000",
            "999999999999999999999",
            "0x0",
            "0xC8",
            "0xC8, 0xFA0",
            "Mouse=0x0, Keyboard=0xC8, Audio=0xFA0",
            "Mouse=invalid, Keyboard=0xC8",
            "Мышь=0x0, Клавиатура=0xC8",
            "0x0, 0xC8, 0xC8, 0xC8, 0xC8, 0xC8, 0xC8, 0xC8, 0xC8, 0xC8",
            "=0x0",
            "Mouse=",
            ",,,",
        ];

        foreach (string input in inputs)
        {
            object?[] args = [input, 0xC8u, null];
            _ = tryParse.Invoke(null, args);
        }
    }

    private static void ConfigureForm(MainForm form, CpuProfileSpec cpu, DeviceInfo[] devices, Size windowSize, float scale)
    {
        typeof(MainForm).GetField("_uiScale", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, scale);
        typeof(MainForm).GetField("_cpuInfo", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, cpu.Cpu);
        typeof(MainForm).GetField("_maxLogical", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, cpu.MaxLogical);

        form.Size = windowSize;
        form.CreateControl();
        form.PerformLayout();

        MethodInfo normalize = typeof(MainForm).GetMethod("NormalizeInstanceId", BindingFlags.Static | BindingFlags.NonPublic)!;
        var powerCache = (Dictionary<string, bool>)typeof(MainForm).GetField("_displayPowerCache", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        var rssCache = (Dictionary<string, NdisRssRuntimeState>)typeof(MainForm).GetField("_ndisRssRuntimeCache", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        powerCache.Clear();
        rssCache.Clear();

        foreach (DeviceInfo device in devices)
        {
            string norm = (string)normalize.Invoke(null, [device.InstanceId])!;
            powerCache[norm] = true;
            if (device.Kind == DeviceKind.NET_NDIS && !device.Wifi)
            {
                rssCache[norm] = new NdisRssRuntimeState(true, true, true, 0, 0, 0, Math.Min(7, cpu.MaxLogical - 1), 4, 4, "Ethernet", device.Name, "1", string.Empty);
            }
        }

        MethodInfo newBlock = typeof(MainForm).GetMethod("NewDeviceBlock", BindingFlags.Instance | BindingFlags.NonPublic)!;
        for (int i = 0; i < devices.Length; i++)
        {
            newBlock.Invoke(form, [devices[i], i, null, null]);
        }

        MethodInfo newReserved = typeof(MainForm).GetMethod("NewReservedCpuSetsPanel", BindingFlags.Instance | BindingFlags.NonPublic)!;
        Panel? reserved = (Panel?)newReserved.Invoke(form, [null]);
        if (reserved is not null)
        {
            var devicesPanel = (Panel)typeof(MainForm).GetField("_devicesPanel", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
            devicesPanel.Controls.Add(reserved);
            typeof(MainForm).GetField("_reservedCpuPanel", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, reserved);
        }

        InvokeLayoutBlocks(form);
    }

    private static void VerifyAllBlocks(MainForm form, string context, List<string> failures)
    {
        MethodInfo logLayout = typeof(MainForm).GetMethod("LogGuiBlockLayout", BindingFlags.Instance | BindingFlags.NonPublic)!;
        List<DeviceBlock> blocks = GetBlocks(form);

        for (int i = 0; i < blocks.Count; i++)
        {
            DeviceBlock block = blocks[i];
            int issues = (int)logLayout.Invoke(form, [block, i])!;
            if (issues > 0)
            {
                string detail = ReadLastLayoutIssue(form);
                failures.Add($"{context} block[{i}] ({block.Kind} {block.Device.Name}): {detail}");
            }

            if (block.SettingsPanel is Panel settings)
            {
                foreach (Control child in settings.Controls)
                {
                    if (!child.Visible) continue;
                    if (child.Right > settings.ClientSize.Width + 2 || child.Bottom > settings.ClientSize.Height + 2)
                    {
                        failures.Add($"{context} block[{i}] child {child.GetType().Name} '{child.Text}' overflows settingsPanel: child={child.Bounds}, panel={settings.ClientSize}");
                    }
                }
            }
        }
    }

    private static void VerifyCardInteractions(MainForm form, DeviceBlock block, string context, List<string> failures)
    {
        VerifyFeedbackStability(form, block, context, failures);

        if (block.ImodCheckButton is not null && block.ImodModeCombo is not null)
        {
            bool initialChecked = block.ImodAutoCheck.Checked;
            int initialMode = block.ImodModeCombo.SelectedIndex;

            block.ImodAutoCheck.Checked = !initialChecked;
            block.RelayoutAction?.Invoke();
            VerifySingleBlock(form, block, $"{context}/imod-toggle", failures);

            for (int m = 0; m < block.ImodModeCombo.Items.Count; m++)
            {
                block.ImodModeCombo.SelectedIndex = m;
                block.RelayoutAction?.Invoke();
                VerifySingleBlock(form, block, $"{context}/imod-mode-{m}", failures);
            }

            if (block.ImodDetailsButton is { Visible: true } detailsButton)
            {
                block.RestoreImodDetailsExpansion?.Invoke(true);
                block.RelayoutAction?.Invoke();
                VerifySingleBlock(form, block, $"{context}/imod-details-open", failures);

                block.RestoreImodDetailsExpansion?.Invoke(false);
                block.RelayoutAction?.Invoke();
            }

            if (initialMode >= 0 && initialMode < block.ImodModeCombo.Items.Count)
            {
                block.ImodModeCombo.SelectedIndex = initialMode;
            }
            block.ImodAutoCheck.Checked = initialChecked;
            block.RelayoutAction?.Invoke();
        }
    }

    private static void VerifySingleBlock(MainForm form, DeviceBlock block, string context, List<string> failures)
    {
        MethodInfo logLayout = typeof(MainForm).GetMethod("LogGuiBlockLayout", BindingFlags.Instance | BindingFlags.NonPublic)!;
        int issues = (int)logLayout.Invoke(form, [block, 0])!;
        if (issues > 0)
        {
            failures.Add($"{context} ({block.Kind} {block.Device.Name}): {ReadLastLayoutIssue(form)}");
        }
    }

    private static void VerifyFeedbackStability(MainForm form, DeviceBlock block, string context, List<string> failures)
    {
        MethodInfo beginFeedback = typeof(MainForm).GetMethod("BeginButtonFeedback", BindingFlags.Instance | BindingFlags.NonPublic)!;
        List<(string Name, Button Button, string Caption, Control Editor)> targets = [];

        if (block.ImodCheckButton is not null && block.SettingsPanel is not null)
        {
            Button? setBtn = block.SettingsPanel.Controls.OfType<Button>().FirstOrDefault(b => GetSourceText(form, b) == "SET");
            Button? delBtn = block.SettingsPanel.Controls.OfType<Button>().FirstOrDefault(b => GetSourceText(form, b) == "DELETE PROFILE");
            Control editor = block.ImodBox.Parent!;
            targets.Add(("usb-check", block.ImodCheckButton, "CHECK…", editor));
            if (delBtn is not null) targets.Add(("usb-delete", delBtn, "DELETE…", editor));
            if (setBtn is not null) targets.Add(("usb-set", setBtn, "SET…", editor));
        }

        if (block.NicItrCheckButton is not null && block.NicItrBox?.Parent is Control nicEditor)
        {
            targets.Add(("nic-check", block.NicItrCheckButton, "CHECK…", nicEditor));
            if (block.NicItrSaveButton is not null) targets.Add(("nic-save", block.NicItrSaveButton, "SAVE…", nicEditor));
            if (block.NicItrApplyButton is not null) targets.Add(("nic-set", block.NicItrApplyButton, "SET…", nicEditor));
        }

        foreach (var (name, button, caption, editor) in targets)
        {
            Rectangle btnBefore = button.Bounds;
            Rectangle edBefore = editor.Bounds;
            using (var scope = (IDisposable)beginFeedback.Invoke(form, [button, caption, null])!)
            {
                block.RelayoutAction?.Invoke();
                if (button.Bounds != btnBefore || editor.Bounds != edBefore)
                {
                    failures.Add($"{context}/{name} drifted during busy feedback: btn {btnBefore}->{button.Bounds}, ed {edBefore}->{editor.Bounds}");
                }
            }

            if (button.Bounds != btnBefore || editor.Bounds != edBefore)
            {
                failures.Add($"{context}/{name} drifted after busy feedback dispose: btn {btnBefore}->{button.Bounds}, ed {edBefore}->{editor.Bounds}");
            }

            if (block.SettingsPanel is Panel panel && button.Right > panel.ClientSize.Width)
            {
                failures.Add($"{context}/{name} clipped after feedback: right={button.Right} > panel={panel.ClientSize.Width}");
            }
        }
    }

    private static HardwareProfileSpec[] BuildHardwareProfiles() =>
    [
        new("laptop-4c8t", BuildLaptop4C8T(),
        [
            MakeUsb("Intel(R) USB 3.10 eXtensible Host Controller - 1.10", @"PCI\VEN_8086&DEV_9A13\3&11583659&0&A0", "Mouse 1K, Keyboard"),
            MakeGpu("Intel(R) Iris(R) Xe Graphics", @"PCI\VEN_8086&DEV_9A49\3&11583659&0&10", integrated: true),
            MakeNic("Intel(R) Wi-Fi 6E AX211 160MHz", @"PCI\VEN_8086&DEV_51F0\3&11583659&0&A3", DeviceKind.NET_NDIS, wifi: true),
            MakeStorage("Standard NVM Express Controller", @"PCI\VEN_144D&DEV_A80A\4&123456&0&00E0", "NVMe"),
        ]),
        new("midrange-6c12t", BuildMidrange6C12T(),
        [
            MakeUsb("AMD USB 3.10 eXtensible Host Controller - 1.10", @"PCI\VEN_1022&DEV_149C\4&2BA2BF07&0&0010", "Mouse 8K, Keyboard 8K, Audio"),
            MakeGpu("NVIDIA GeForce RTX 4070 SUPER", @"PCI\VEN_10DE&DEV_2783\4&2BA2BF07&0&0008"),
            MakeNic("Realtek Gaming 2.5GbE Family Controller", @"PCI\VEN_10EC&DEV_8125\4&17112BEF&0&00E2", DeviceKind.NET_NDIS),
            MakeAudio("High Definition Audio Controller", @"PCI\VEN_1022&DEV_1487\4&2BA2BF07&0&0020", "Speakers"),
            MakeStorage("Standard SATA AHCI Controller", @"PCI\VEN_1022&DEV_7901\4&111111&0&0000", "SSD"),
        ]),
        Build8Core16ThreadGamingProfile(),
        new("hybrid-14600k-20t", BuildHybrid14600K(),
        [
            MakeUsb("Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7A60\3&11583659&0&A0", "Mouse 4K, Keyboard 1K, Gamepad"),
            MakeGpu("NVIDIA GeForce RTX 4070 Ti", @"PCI\VEN_10DE&DEV_2782\4&2BA2BF07&0&0008"),
            MakeNic("Intel(R) Ethernet Controller I225-V", @"PCI\VEN_8086&DEV_15F3\4&17112BEF&0&00E2", DeviceKind.NET_NDIS),
            MakeStorage("Standard NVM Express Controller", @"PCI\VEN_1CC1&DEV_622A\4&97A1985&0&0030", "NVMe"),
        ]),
        BuildHybrid14900KProfile(),
        BuildDualCcd9950X3DProfile(),
        new("ultra-285k-24t-nosmt", BuildCoreUltra285KNoSmt(),
        [
            MakeUsb("Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7F6E\3&11583659&0&A0", "Mouse 8K, Keyboard 8K"),
            MakeGpu("NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\4&2BA2BF07&0&0008"),
            MakeNic("Intel(R) Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\4&17112BEF&0&00E2", DeviceKind.NET_NDIS),
            MakeNic("Intel(R) Wi-Fi 7 BE200 320MHz", @"PCI\VEN_8086&DEV_272B\4&17112BEF&0&00E4", DeviceKind.NET_NDIS, wifi: true),
        ]),
        new("server-64t", BuildServer64Thread(),
        [
            MakeUsb("AMD USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_15B6\4&2BA2BF07&0&0010", "Mouse 1K, Keyboard 1K, Audio, Microphone, Gamepad"),
            MakeGpu("AMD Radeon RX 9070 XT", @"PCI\VEN_1002&DEV_7550\4&2BA2BF07&0&0008"),
            MakeNic("Intel(R) Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\4&17112BEF&0&00E2", DeviceKind.NET_CX),
            MakeStorage("Standard NVM Express Controller", @"PCI\VEN_1CC1&DEV_622A\4&97A1985&0&0030", "NVMe"),
        ]),
    ];

    private static HardwareProfileSpec Build8Core16ThreadGamingProfile() => new("gaming-8c16t", BuildGaming8C16T(),
    [
        MakeUsb("Расширяемый хост-контроллер Intel(R) USB 3.20 — 1.20", @"PCI\VEN_8086&DEV_7AE0&SUBSYS_7D361462&REV_11\3&11583659&0&A0", "Mouse 8K, Keyboard 8K, Microphone, Audio"),
        MakeGpu("NVIDIA GeForce RTX 5060 Ti", @"PCI\VEN_10DE&DEV_2D04&SUBSYS_53621462&REV_A1\4&2BA2BF07&0&0008"),
        MakeNic("Realtek Gaming 2.5GbE Family Controller", @"PCI\VEN_10EC&DEV_8125&SUBSYS_7D361462&REV_05\4&17112BEF&0&00E2", DeviceKind.NET_NDIS),
        MakeAudio("NVIDIA High Definition Audio", @"PCI\VEN_10DE&DEV_22EB\4&2BA2BF07&0&0009", "DisplayPort Audio"),
        MakeStorage("Standard NVM Express Controller", @"PCI\VEN_1CC1&DEV_622A\4&97A1985&0&0030", "NVMe"),
    ]);

    private static HardwareProfileSpec BuildHybrid14900KProfile() => new("hybrid-14900k-32t", BuildHybrid14900K(),
    [
        MakeUsb("Intel(R) USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_8086&DEV_7A60\3&11583659&0&A0", "Mouse 8K, Keyboard 8K, Audio"),
        MakeGpu("NVIDIA GeForce RTX 4090", @"PCI\VEN_10DE&DEV_2684\4&2BA2BF07&0&0008"),
        MakeNic("Intel(R) Ethernet Controller I226-V", @"PCI\VEN_8086&DEV_125C\4&17112BEF&0&00E2", DeviceKind.NET_NDIS),
        MakeStorage("Standard NVM Express Controller", @"PCI\VEN_144D&DEV_A80A\4&97A1985&0&0030", "NVMe"),
    ]);

    private static HardwareProfileSpec BuildDualCcd9950X3DProfile() => new("dualccd-9950x3d-32t", BuildDualCcd9950X3D(),
    [
        MakeUsb("AMD USB 3.20 eXtensible Host Controller - 1.20", @"PCI\VEN_1022&DEV_15B6\4&2BA2BF07&0&0010", "Mouse 8K, Keyboard 8K, Microphone, Audio, Gamepad"),
        MakeGpu("NVIDIA GeForce RTX 5090", @"PCI\VEN_10DE&DEV_2B85\4&2BA2BF07&0&0008"),
        MakeNic("Realtek Gaming 5GbE Family Controller", @"PCI\VEN_10EC&DEV_8126\4&17112BEF&0&00E2", DeviceKind.NET_CX),
        MakeStorage("Standard NVM Express Controller", @"PCI\VEN_144D&DEV_A80A\4&97A1985&0&0030", "NVMe"),
    ]);

    private static CpuProfileSpec BuildLaptop4C8T() => MakeUniformCpu("Laptop 4c/8t", cores: 4, smt: true);
    private static CpuProfileSpec BuildMidrange6C12T() => MakeUniformCpu("Ryzen 6c/12t", cores: 6, smt: true);
    private static CpuProfileSpec BuildGaming8C16T() => MakeUniformCpu("Gaming 8c/16t", cores: 8, smt: true);
    private static CpuProfileSpec BuildServer64Thread() => MakeUniformCpu("Server 32c/64t", cores: 32, smt: true);

    private static CpuProfileSpec BuildHybrid14600K() => MakeHybridCpu("Intel 14600K 6P+8E", pCores: 6, pSmt: true, eCores: 8);
    private static CpuProfileSpec BuildHybrid14900K() => MakeHybridCpu("Intel 14900K 8P+16E", pCores: 8, pSmt: true, eCores: 16);
    private static CpuProfileSpec BuildCoreUltra285KNoSmt() => MakeHybridCpu("Core Ultra 285K 8P+16E", pCores: 8, pSmt: false, eCores: 16);

    private static CpuProfileSpec BuildDualCcd9950X3D()
    {
        List<CpuLpInfo> lps = [];
        Dictionary<int, int> ccd = [];
        Dictionary<int, int> ccx = [];
        for (int lp = 0; lp < 32; lp++)
        {
            int ccdIndex = lp < 16 ? 0 : 1;
            long l3 = ccdIndex == 0 ? 96L * 1024 * 1024 : 32L * 1024 * 1024;
            lps.Add(new CpuLpInfo(0, lp, lp / 2, ccdIndex, 0, 0, lp, -1, l3));
            ccd[lp] = ccdIndex;
            ccx[lp] = ccdIndex;
        }

        return new("Ryzen 9 9950X3D", new CpuInfo
        {
            Topology = new CpuTopology(lps),
            CcdMap = ccd,
            CcxMap = ccx,
            CcdL3CacheBytes = new Dictionary<int, long> { [0] = 96L * 1024 * 1024, [1] = 32L * 1024 * 1024 },
        }, 32, HasHybridCores: false, HasDualCcd: true);
    }

    private static CpuProfileSpec MakeUniformCpu(string name, int cores, bool smt)
    {
        int threadsPerCore = smt ? 2 : 1;
        int logical = cores * threadsPerCore;
        List<CpuLpInfo> lps = [];
        Dictionary<int, int> ccd = [];
        Dictionary<int, int> ccx = [];
        for (int lp = 0; lp < logical; lp++)
        {
            lps.Add(new CpuLpInfo(0, lp, lp / threadsPerCore, 0, 0, 0, lp));
            ccd[lp] = 0;
            ccx[lp] = 0;
        }

        return new(name, new CpuInfo
        {
            Topology = new CpuTopology(lps),
            CcdMap = ccd,
            CcxMap = ccx,
            CcdL3CacheBytes = new Dictionary<int, long>(),
        }, logical, HasHybridCores: false, HasDualCcd: false);
    }

    private static CpuProfileSpec MakeHybridCpu(string name, int pCores, bool pSmt, int eCores)
    {
        List<CpuLpInfo> lps = [];
        Dictionary<int, int> ccd = [];
        Dictionary<int, int> ccx = [];
        int lp = 0;
        int coreId = 0;
        for (int p = 0; p < pCores; p++)
        {
            int count = pSmt ? 2 : 1;
            for (int t = 0; t < count; t++)
            {
                lps.Add(new CpuLpInfo(0, lp, coreId, 0, 0, EffClass: 1, LocalIndex: lp));
                ccd[lp] = 0;
                ccx[lp] = 0;
                lp++;
            }
            coreId++;
        }

        for (int e = 0; e < eCores; e++)
        {
            lps.Add(new CpuLpInfo(0, lp, coreId, 0, 0, EffClass: 0, LocalIndex: lp));
            ccd[lp] = 0;
            ccx[lp] = 0;
            lp++;
            coreId++;
        }

        return new(name, new CpuInfo
        {
            Topology = new CpuTopology(lps),
            CcdMap = ccd,
            CcxMap = ccx,
            CcdL3CacheBytes = new Dictionary<int, long>(),
        }, lp, HasHybridCores: true, HasDualCcd: false);
    }

    private static List<AutoCpuUnit> BuildAutoCpuUnits(CpuProfileSpec cpu)
    {
        return cpu.Cpu.Topology.ByCore
            .OrderBy(kv => kv.Value.Min(lp => lp.LP))
            .Select((kv, idx) =>
            {
                List<int> coreLps = kv.Value.Select(x => x.LP).OrderBy(x => x).ToList();
                int primary = coreLps[0];
                bool isECore = cpu.HasHybridCores && kv.Value[0].EffClass == 0;
                return new AutoCpuUnit
                {
                    Id = idx,
                    PrimaryLp = primary,
                    Lps = coreLps,
                    IsECore = isECore,
                    Ccd = cpu.Cpu.CcdMap.GetValueOrDefault(primary, 0),
                    Ccx = cpu.Cpu.CcxMap.GetValueOrDefault(primary, 0),
                    Rating = isECore ? 100 : 150 - idx,
                    Rank = idx + 1,
                };
            })
            .ToList();
    }

    private static DeviceInfo MakeUsb(string name, string id, string roles) => new()
    {
        Name = name,
        InstanceId = id,
        Class = "USB",
        RegBase = string.Empty,
        Kind = DeviceKind.USB,
        UsbRoles = roles,
        UsbIsXhci = true,
        UsbHasDevices = true,
    };

    private static DeviceInfo MakeGpu(string name, string id, bool integrated = false) => new()
    {
        Name = name,
        InstanceId = id,
        Class = "Display",
        RegBase = string.Empty,
        Kind = DeviceKind.GPU,
        IsIntegratedGpu = integrated,
    };

    private static DeviceInfo MakeNic(string name, string id, DeviceKind kind, bool wifi = false) => new()
    {
        Name = name,
        InstanceId = id,
        Class = "Net",
        RegBase = string.Empty,
        Kind = kind,
        Wifi = wifi,
    };

    private static DeviceInfo MakeAudio(string name, string id, string endpoints) => new()
    {
        Name = name,
        InstanceId = id,
        Class = "MEDIA",
        RegBase = string.Empty,
        Kind = DeviceKind.AUDIO,
        AudioEndpoints = endpoints,
    };

    private static DeviceInfo MakeStorage(string name, string id, string tag) => new()
    {
        Name = name,
        InstanceId = id,
        Class = "SCSIAdapter",
        RegBase = string.Empty,
        Kind = DeviceKind.STOR,
        StorageTag = tag,
    };

    private static void InvokeLayoutBlocks(MainForm form)
    {
        MethodInfo layoutBlocks = typeof(MainForm).GetMethod("LayoutBlocks", BindingFlags.Instance | BindingFlags.NonPublic)!;
        layoutBlocks.Invoke(form, [true]);

        // Hidden test forms have Control.Visible == false, so LayoutBlocks skips
        // per-card width/relayout updates; apply the exact same card pass here.
        MethodInfo getViewport = typeof(MainForm).GetMethod("GetDevicesViewportWidth", BindingFlags.Instance | BindingFlags.NonPublic)!;
        MethodInfo uiScale = typeof(MainForm).GetMethod("UiScale", BindingFlags.Instance | BindingFlags.NonPublic, [typeof(int)])!;
        int cardWidth = Math.Max((int)uiScale.Invoke(form, [360])!, (int)getViewport.Invoke(form, null)! - (int)uiScale.Invoke(form, [36])!);
        foreach (DeviceBlock block in GetBlocks(form))
        {
            block.Group.Width = cardWidth;
            block.RelayoutAction?.Invoke();
        }
    }

    private static List<DeviceBlock> GetBlocks(MainForm form)
    {
        object? value = typeof(MainForm).GetField("_blocks", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form);
        return Assert.IsType<List<DeviceBlock>>(value);
    }

    private static string ReadLastLayoutIssue(MainForm form)
    {
        string? logPath = AppDiagnostics.SessionLogPath;
        if (string.IsNullOrWhiteSpace(logPath) || !File.Exists(logPath))
        {
            return "LogGuiBlockLayout reported issue (no log file)";
        }

        List<string> lines = [];
        using (FileStream fs = new(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (StreamReader reader = new(fs))
        {
            while (reader.ReadLine() is string line) lines.Add(line);
        }
        int idx = lines.FindLastIndex(line => line.Contains("GUI.LAYOUT.ISSUE", StringComparison.Ordinal));
        if (idx < 0)
        {
            return "LogGuiBlockLayout reported issue (no GUI.LAYOUT.ISSUE line)";
        }

        List<string> snippet = [lines[idx]];
        for (int back = Math.Max(0, idx - 3); back < idx; back++)
        {
            if (lines[back].Contains("GUI.LAYOUT.SETTINGS-", StringComparison.Ordinal))
            {
                snippet.Insert(0, lines[back]);
            }
        }
        return string.Join(" | ", snippet);
    }

    private static string GetSourceText(MainForm form, Control control)
    {
        MethodInfo method = typeof(MainForm).GetMethod("GetSourceControlText", BindingFlags.Instance | BindingFlags.NonPublic)!;
        return (string)method.Invoke(form, [control])!;
    }

    private static void RunSta(Action action)
    {
        Exception? failure = null;
        Thread thread = new(() =>
        {
            try { action(); }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(60000), "STA hardware matrix thread timed out.");
        if (failure is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
