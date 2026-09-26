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
        else if (showcase.StartsWith("Imod", StringComparison.OrdinalIgnoreCase))
        {
            SetCategoryFilter("USB");
        }

        UpdateAllBlocksInitialState();
        UpdateApplyButtonDirtyCount();
        CloseDevicesBusyOverlay();
        if (_devicesBusyOverlay is not null)
        {
            _devicesBusyOverlay.Visible = false;
        }
        SetOperationButtonsEnabled(true);
        if (_devicesScroll is not null)
        {
            _devicesScroll.Value = 0;
        }
        _searchFilterBox?.Inner.Select(0, 0);
        _btnScanRef?.Focus();
        ActiveControl = _btnScanRef;

        WriteLog($"SHOWCASE: initialized mode={showcase} language={(isRu ? "ru" : "en")} blocks={_blocks.Count}");
        return true;
    }

    private void SetupRyzen9950X3DShowcase(bool isRu)
    {
        // 16 cores, 32 threads, 2 CCDs (8 cores CCD0 with 3D V-Cache, 8 cores CCD1 frequency)
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

        // GPU: NVIDIA RTX 5090
        _testDevices.Add(CreateTestDevice(
            DeviceKind.GPU,
            "NVIDIA GeForce RTX 5090",
            @"PCI\VEN_10DE&DEV_2B85&SUBSYS_14620000\4&2BA2BF07&0&0008",
            usbRoles: "", audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: false, usbHasDevices: false, integratedGpu: false,
            testIrqCount: 1, testMsiStatus: "Enabled"));

        // USB 1 (CPU-direct, CHIP 0): Dedicated to Mouse 8K
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_1022&DEV_15B6&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K" : "Mouse 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // USB 2 (Chipset, CHIP 1): Dedicated to Keyboard 8K
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.10 eXtensible Host Controller - 1.10",
            @"PCI\VEN_1022&DEV_43F7&SUBSYS_14620000\4&2BA2BF07&0&0012",
            usbRoles: isRu ? "Клавиатура 8K" : "Keyboard 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // USB 3 (ASMedia Addon, CHIP 1+): Dedicated to Audio DAC & Microphone
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "ASMedia USB 3.1 eXtensible Host Controller - 1.10",
            @"PCI\VEN_1B21&DEV_1242&SUBSYS_14620000\4&2BA2BF07&0&0014",
            usbRoles: isRu ? "Аудио ЦАП, Микрофон" : "Audio DAC, Microphone",
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

        // Storage: Samsung 990 PRO NVMe
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
                    int targetLp = block.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                        || block.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase)
                        ? 6
                        : (block.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                            || block.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase)
                            ? 8
                            : 10);

                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == targetLp;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PolicyCombo.SelectedItem = "SpecCPU";

                if (block.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                    || block.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
                {
                    block.PrioCombo.SelectedItem = "High";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0x0";
                    if (block.RawMouseThrottleCheck is not null)
                    {
                        block.RawMouseThrottleCheck.Enabled = true;
                        block.RawMouseThrottleCheck.Checked = true;
                    }
                    if (block.RawMouseThrottleCombo is not null)
                    {
                        block.RawMouseThrottleCombo.Enabled = true;
                        block.RawMouseThrottleCombo.SelectedItem = GetRawMouseThrottlePreset(20);
                    }
                    if (block.RawMouseThrottleStatusLabel is not null)
                    {
                        block.RawMouseThrottleStatusLabel.Text = "current: 50Hz (DWORD=20)";
                        block.RawMouseThrottleStatusLabel.ForeColor = _statusActive;
                    }
                    RefreshTestImodPreview(block, "test-imod-set");
                }
                else if (block.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                    || block.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
                {
                    block.PrioCombo.SelectedItem = "High";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0x0";
                    RefreshTestImodPreview(block, "test-imod-set");
                }
                else
                {
                    block.PrioCombo.SelectedItem = "Normal";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0xFA0";
                    RefreshTestImodPreview(block, "test-imod-set");
                }

                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to CCD0 physical core 12 (LP 12)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 12;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
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

        // USB 1 (CPU-direct, CHIP 0): Dedicated to Mouse 8K
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "Intel(R) USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_8086&DEV_461E&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K" : "Mouse 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // USB 2 (Chipset PCH, CHIP 1): Dedicated to Keyboard 8K
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "Intel(R) USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_8086&DEV_7A60&SUBSYS_14620000\4&2BA2BF07&0&0012",
            usbRoles: isRu ? "Клавиатура 8K" : "Keyboard 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // USB 3 (ASMedia Addon, CHIP 1+): Dedicated to Audio DAC & Microphone
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "ASMedia USB 3.1 eXtensible Host Controller - 1.10",
            @"PCI\VEN_1B21&DEV_1242&SUBSYS_14620000\4&2BA2BF07&0&0014",
            usbRoles: isRu ? "Аудио ЦАП, Микрофон" : "Audio DAC, Microphone",
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

        // Storage: Samsung 990 PRO NVMe
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
                    int targetLp = block.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                        || block.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase)
                        ? 6
                        : (block.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                            || block.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase)
                            ? 8
                            : 10);

                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == targetLp;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.LimitBox.Text = "0";
                block.PolicyCombo.SelectedItem = "SpecCPU";

                if (block.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                    || block.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
                {
                    block.PrioCombo.SelectedItem = "High";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0x0";
                    if (block.RawMouseThrottleCheck is not null)
                    {
                        block.RawMouseThrottleCheck.Enabled = true;
                        block.RawMouseThrottleCheck.Checked = true;
                    }
                    if (block.RawMouseThrottleCombo is not null)
                    {
                        block.RawMouseThrottleCombo.Enabled = true;
                        block.RawMouseThrottleCombo.SelectedItem = GetRawMouseThrottlePreset(20);
                    }
                    if (block.RawMouseThrottleStatusLabel is not null)
                    {
                        block.RawMouseThrottleStatusLabel.Text = "current: 50Hz (DWORD=20)";
                        block.RawMouseThrottleStatusLabel.ForeColor = _statusActive;
                    }
                    RefreshTestImodPreview(block, "test-imod-set");
                }
                else if (block.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                    || block.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
                {
                    block.PrioCombo.SelectedItem = "High";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0x0";
                    RefreshTestImodPreview(block, "test-imod-set");
                }
                else
                {
                    block.PrioCombo.SelectedItem = "Normal";
                    block.ImodAutoCheck.Checked = false;
                    block.ImodBox.Text = "0xFA0";
                    RefreshTestImodPreview(block, "test-imod-set");
                }

                RecalcAffinityMask(block);
                OnBlockSettingChanged(block);
            }
            else if (block.Kind == DeviceKind.NET_NDIS)
            {
                block.SuppressCpuEvents++;
                try
                {
                    // Assign to P-Core 12 (LP 12)
                    for (int i = 0; i < block.CpuBoxes.Count; i++)
                    {
                        block.CpuBoxes[i].Checked = i == 12;
                    }
                }
                finally { block.SuppressCpuEvents--; }

                block.MsiCombo.SelectedItem = "Enabled";
                block.PrioCombo.SelectedItem = "High";
                block.PolicyCombo.SelectedItem = "SpecCPU";
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

        // Controller 1: Dedicated Mouse 8K Controller (CPU-Direct, CHIP 0)
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.20 eXtensible Host Controller - 1.20",
            @"PCI\VEN_1022&DEV_15B6&SUBSYS_14620000\4&2BA2BF07&0&0010",
            usbRoles: isRu ? "Мышь 8K" : "Mouse 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // Controller 2: Dedicated Keyboard 8K Controller (Chipset, CHIP 1)
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "AMD USB 3.10 eXtensible Host Controller - 1.10",
            @"PCI\VEN_1022&DEV_43F7&SUBSYS_14620000\4&2BA2BF07&0&0012",
            usbRoles: isRu ? "Клавиатура 8K" : "Keyboard 8K",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        // Controller 3: Dedicated Audio Controller (ASMedia Addon, CHIP 1+)
        _testDevices.Add(CreateTestDevice(
            DeviceKind.USB,
            "ASMedia USB 3.1 eXtensible Host Controller - 1.10",
            @"PCI\VEN_1B21&DEV_1242&SUBSYS_14620000\4&2BA2BF07&0&0014",
            usbRoles: isRu ? "Аудио ЦАП, Микрофон" : "Audio DAC, Microphone",
            audioEndpoints: "", storageTag: "", wifi: false,
            usbIsXhci: true, usbHasDevices: true, integratedGpu: false,
            testIrqCount: 2, testMsiStatus: "Enabled", usbSelectiveSuspend: "off"));

        RefreshBlocks();

        foreach (DeviceBlock usbBlock in _blocks.Where(b => b.Kind == DeviceKind.USB))
        {
            usbBlock.SuppressCpuEvents++;
            try
            {
                int targetLp = usbBlock.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                    || usbBlock.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase)
                    ? 6
                    : (usbBlock.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                        || usbBlock.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase)
                        ? 8
                        : 10);

                for (int i = 0; i < usbBlock.CpuBoxes.Count; i++)
                {
                    usbBlock.CpuBoxes[i].Checked = i == targetLp;
                }
            }
            finally { usbBlock.SuppressCpuEvents--; }

            usbBlock.MsiCombo.SelectedItem = "Enabled";
            usbBlock.LimitBox.Text = "0";
            usbBlock.PolicyCombo.SelectedItem = "SpecCPU";
            usbBlock.ImodAutoCheck.Checked = true;

            if (usbBlock.Device.UsbRoles.Contains("Мышь", StringComparison.OrdinalIgnoreCase)
                || usbBlock.Device.UsbRoles.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
            {
                usbBlock.PrioCombo.SelectedItem = "High";
                usbBlock.ImodBox.Text = "0x0";
                if (usbBlock.RawMouseThrottleCheck is not null)
                {
                    usbBlock.RawMouseThrottleCheck.Enabled = true;
                    usbBlock.RawMouseThrottleCheck.Checked = true;
                }
                if (usbBlock.RawMouseThrottleCombo is not null)
                {
                    usbBlock.RawMouseThrottleCombo.Enabled = true;
                    usbBlock.RawMouseThrottleCombo.SelectedItem = GetRawMouseThrottlePreset(20);
                }
                if (usbBlock.RawMouseThrottleStatusLabel is not null)
                {
                    usbBlock.RawMouseThrottleStatusLabel.Text = "current: 50Hz (DWORD=20)";
                    usbBlock.RawMouseThrottleStatusLabel.ForeColor = _statusActive;
                }
                RefreshTestImodPreview(usbBlock, "test-imod-set");
            }
            else if (usbBlock.Device.UsbRoles.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase)
                || usbBlock.Device.UsbRoles.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
            {
                usbBlock.PrioCombo.SelectedItem = "High";
                usbBlock.ImodBox.Text = "0x0";
                RefreshTestImodPreview(usbBlock, "test-imod-set");
            }
            else
            {
                usbBlock.PrioCombo.SelectedItem = "Normal";
                usbBlock.ImodBox.Text = "0xFA0";
                RefreshTestImodPreview(usbBlock, "test-imod-set");
            }

            RecalcAffinityMask(usbBlock);
            OnBlockSettingChanged(usbBlock);
        }

        UpdateAllBlocksInitialState();
        UpdateApplyButtonDirtyCount();
        UpdateCpuHeaderUi();
    }
}
