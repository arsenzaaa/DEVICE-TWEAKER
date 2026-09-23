using System.Drawing;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool CheckAndApplyShowcaseSetup()
    {
        string? showcase = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_SHOWCASE");
        if (string.IsNullOrWhiteSpace(showcase))
        {
            return false;
        }

        Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_HIDE_SANDBOX_HEADER", "1");

        bool isRu = showcase.EndsWith("_RU", StringComparison.OrdinalIgnoreCase);
        UiLanguage.Set(isRu ? UiLanguageCode.Russian : UiLanguageCode.English, persist: false);

        _testDevicesEnabled = true;
        _testDevicesOnly = true;
        _testAutoDryRun = true;

        if (showcase.StartsWith("Ryzen9950X3D", StringComparison.OrdinalIgnoreCase))
        {
            SetupRyzen9950X3DShowcase(isRu);
        }
        else if (showcase.StartsWith("Intel14900K", StringComparison.OrdinalIgnoreCase))
        {
            SetupIntel14900KShowcase(isRu);
        }
        else if (showcase.StartsWith("Imod", StringComparison.OrdinalIgnoreCase))
        {
            SetupImodShowcase(isRu);
        }
        else
        {
            return false;
        }

        string? filter = Environment.GetEnvironmentVariable("DEVICE_TWEAKER_SHOWCASE_FILTER");
        if (!string.IsNullOrWhiteSpace(filter))
        {
            SetCategoryFilter(filter);
        }

        UpdateAllBlocksInitialState();
        UpdateApplyButtonDirtyCount();
        if (_devicesBusyOverlay is not null)
        {
            _devicesBusyOverlay.Visible = false;
        }
        _devicesBusyDepth = 0;
        _devicesBusyDone = 0;
        SetOperationButtonsEnabled(true);
        if (_devicesScroll is not null)
        {
            _devicesScroll.Value = 0;
        }
        _devicesPanel.Location = new Point(0, 0);
        _devicesPanel.Invalidate(true);
        _devicesHost.Invalidate(true);
        Invalidate(true);

        WriteLog($"SHOWCASE: initialized mode={showcase} language={(isRu ? "ru" : "en")} blocks={_blocks.Count}");
        return true;
    }

    private void SetupRyzen9950X3DShowcase(bool isRu)
    {
        // 16 cores, 32 threads, 2 CCDs (8 cores CCD0, 8 cores CCD1)
        TestCpuConfig config = new()
        {
            LogicalCount = 32,
            SmtEnabled = true,
            UseHyperThreadingLabel = false,
            CpuName = "AMD Ryzen 9 9950X3D",
            CcdMap = new Dictionary<int, int>(),
            CcxMap = new Dictionary<int, int>()
        };

        for (int lp = 0; lp < 32; lp++)
        {
            config.CoreMap[lp] = lp / 2;
            int ccd = lp < 16 ? 0 : 1;
            config.CcdMap[lp] = ccd;
            config.CcxMap[lp] = ccd;
            // CCD0 (3D V-Cache): high rating, CCD1: frequency
            config.CppcRatings[lp] = lp < 16 ? 140 - (lp / 2) : 112 - ((lp - 16) / 2);
        }

        ApplyTestCpuConfig(config);

        _testDevices.Clear();

        // GPU: NVIDIA RTX 5090
        _testDevices.Add(CreateTestDevice(
            DeviceKind.GPU,
            "NVIDIA GeForce RTX 5090",
            @"PCI\VEN_10DE&DEV_2B85&SUBSYS_14620000\4&2BA2BF07&0&0008",
            usbRoles: "", audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 1, testMsiStatus: "Enabled"));

        // USB: AMD USB 3.20 (CHIP 0, CPU-direct)
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_1022&DEV_15B6&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K, Клавиатура 8K" : "Mouse 8K, Keyboard 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // Network: Intel I226-V
        _testDevices.Add(CreateTestDevice(
            DeviceKind.NET_NDIS,
            "Intel Ethernet Controller I226-V",
            @"PCI\VEN_8086&DEV_125C&SUBSYS_14620000\4&2BA2BF07&0&0018",
            usbRoles: "", audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 4, testMsiStatus: "Enabled", nicPowerSaving: "off"));

        // Storage: Samsung 990 PRO
        _testDevices.Add(CreateTestDevice(
            DeviceKind.STOR,
            "Samsung 990 PRO NVMe Controller",
            @"PCI\VEN_144D&DEV_A80C&SUBSYS_14620000\4&2BA2BF07&0&0020",
            usbRoles: "", audioEndpoints: "", storageTag: "NVMe", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 1, testMsiStatus: "Enabled"));

        RefreshBlocks();

        // Configure optimal settings on blocks
        foreach (DeviceBlock block in _blocks)
        {
            if (block.Kind == DeviceKind.GPU)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to CCD0 physical cores 2 & 4 (LPs 2, 4)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 2 || i == 4;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.USB)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to CCD0 physical core 6 (LP 6)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 6;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to CCD0 physical core 8 (LP 8)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 8;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                if (block.RssQueueBox is not null) block.RssQueueBox.Value = 4;
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
        }

        UpdateCpuHeaderUi();
    }

    private void SetupIntel14900KShowcase(bool isRu)
    {
        // 8 P-Cores with HT = 16 LPs, 16 E-Cores = 16 LPs -> 32 logical threads
        TestCpuConfig config = new()
        {
            LogicalCount = 32,
            SmtEnabled = true,
            UseHyperThreadingLabel = true,
            CpuName = "Intel Core i9-14900K",
            CcdMap = new Dictionary<int, int>(),
            CcxMap = new Dictionary<int, int>()
        };

        for (int lp = 0; lp < 32; lp++)
        {
            config.CcdMap[lp] = 0;
            config.CcxMap[lp] = 0;
            if (lp < 16)
            {
                config.CoreMap[lp] = lp / 2;
                config.CppcRatings[lp] = 140 - (lp / 2 * 3);
            }
            else
            {
                config.CoreMap[lp] = 8 + (lp - 16);
                config.ECoreLps.Add(lp);
                config.CppcRatings[lp] = 70 - ((lp - 16) / 2);
            }
        }

        ApplyTestCpuConfig(config);

        _testDevices.Clear();

        // GPU: NVIDIA RTX 4090
        _testDevices.Add(CreateTestDevice(
            DeviceKind.GPU,
            "NVIDIA GeForce RTX 4090",
            @"PCI\VEN_10DE&DEV_2684&SUBSYS_14620000\4&2BA2BF07&0&0008",
            usbRoles: "", audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 1, testMsiStatus: "Enabled"));

        // USB: Intel USB 3.20 (CHIP 1, PCH)
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "Intel(R) USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_8086&DEV_7AE0&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K, Клавиатура 8K, Аудио, Микрофон" : "Mouse 8K, Keyboard 8K, Audio, Microphone",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // Network: Intel I226-V
        _testDevices.Add(CreateTestDevice(
            DeviceKind.NET_NDIS,
            "Intel Ethernet Controller I226-V",
            @"PCI\VEN_8086&DEV_125C&SUBSYS_14620000\4&2BA2BF07&0&0018",
            usbRoles: "", audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 4, testMsiStatus: "Enabled", nicPowerSaving: "off"));

        // Storage: Samsung 990 PRO
        _testDevices.Add(CreateTestDevice(
            DeviceKind.STOR,
            "Samsung 990 PRO NVMe Controller",
            @"PCI\VEN_144D&DEV_A80C&SUBSYS_14620000\4&2BA2BF07&0&0020",
            usbRoles: "", audioEndpoints: "", storageTag: "NVMe", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 1, testMsiStatus: "Enabled"));

        RefreshBlocks();

        // Configure optimal settings on blocks (P-Cores only)
        foreach (DeviceBlock block in _blocks)
        {
            if (block.Kind == DeviceKind.GPU)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to P-Cores 2 & 4 (LPs 2, 4)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 2 || i == 4;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.USB)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to P-Core 6 (LP 6)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 6;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to P-Core 8 (LP 8)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 8;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                if (block.RssQueueBox is not null) block.RssQueueBox.Value = 4;
                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
        }

        UpdateCpuHeaderUi();
    }

    private void SetupImodShowcase(bool isRu)
    {
        TestCpuConfig config = new()
        {
            LogicalCount = 32,
            SmtEnabled = true,
            UseHyperThreadingLabel = false,
            CpuName = "AMD Ryzen 9 9950X3D",
            CcdMap = new Dictionary<int, int>(),
            CcxMap = new Dictionary<int, int>()
        };

        for (int lp = 0; lp < 32; lp++)
        {
            config.CoreMap[lp] = lp / 2;
            int ccd = lp < 16 ? 0 : 1;
            config.CcdMap[lp] = ccd;
            config.CcxMap[lp] = ccd;
            config.CppcRatings[lp] = lp < 16 ? 140 - (lp / 2) : 112 - ((lp - 16) / 2);
        }

        ApplyTestCpuConfig(config);

        _testDevices.Clear();

        // USB Controller with IMOD targets
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_1022&DEV_15B6&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K, Клавиатура 8K, Аудио, Микрофон" : "Mouse 8K, Keyboard 8K, Audio, Microphone",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        RefreshBlocks();

        DeviceBlock? usbBlock = _blocks.FirstOrDefault(b => b.Kind == DeviceKind.USB);
        if (usbBlock is not null)
        {
            usbBlock.SuppressCpuEvents++;
            try
            {
                for (int i = 0; i < usbBlock.CpuBoxes.Count; i++)
                {
                    usbBlock.CpuBoxes[i].Checked = i == 4;
                }
            }
            finally { usbBlock.SuppressCpuEvents--; }

            usbBlock.MsiCombo.SelectedItem = "Enabled";
            usbBlock.LimitBox.Text = "0";
            usbBlock.PrioCombo.SelectedItem = "High";
            usbBlock.PolicyCombo.SelectedItem = "SpecCPU";
            usbBlock.ImodAutoCheck.Checked = true;
            usbBlock.ImodBox.Text = "Mouse=0x0, Keyboard=0xC8, Audio=0xFA0";
            if (usbBlock.ImodModeCombo is not null)
            {
                usbBlock.ImodModeCombo.SelectedItem = ImodModeRoles;
            }

            RecalcAffinityMask(usbBlock);
            RefreshTestImodPreview(usbBlock, "test-imod-set");
            OnBlockSettingChanged(usbBlock);
        }

        UpdateAllBlocksInitialState();
        UpdateApplyButtonDirtyCount();
        UpdateCpuHeaderUi();
    }
}
