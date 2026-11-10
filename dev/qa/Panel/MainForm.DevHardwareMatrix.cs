using System.Diagnostics;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private bool _devMatrixSelecting;
    private string[] _devMatrixSystems = [];
    private string[] _devMatrixCpus = [];
    private Func<string, bool>? _devMatrixLoadSystem;
    private Func<string, TestCpuConfig>? _devMatrixLoadCpu;

    private async Task DevAuditHardwareMatrixAsync(string id, string mode, string profileFilter)
    {
        bool presentationOnly = string.Equals(mode, "presentation", StringComparison.Ordinal);
        bool cpuOnly = string.Equals(mode, "cpu", StringComparison.Ordinal);
        if (!presentationOnly && !cpuOnly && !string.IsNullOrEmpty(mode)) throw new InvalidOperationException("Unknown hardware matrix mode: " + mode);
        if (!QaRun.IsReadOnly || !_testDevicesOnly)
            throw new InvalidOperationException("Hardware matrix requires immutable read-only synthetic mode.");
        async Task OpenPanel()
        {
            DevQueue(ShowTestAdminDialog);
            Stopwatch wait = Stopwatch.StartNew();
            while (_devMatrixLoadSystem is null && wait.ElapsedMilliseconds < 10000) await Task.Delay(20);
            if (_devMatrixLoadSystem is null || _devMatrixLoadCpu is null)
                throw new InvalidOperationException("The owned test panel did not initialize its preset loaders.");
        }
        async Task ClosePanel()
        {
            Form panel = DevActiveForm();
            if (panel.Name != "TEST_ADMIN_DIALOG") throw new InvalidOperationException("Unexpected matrix dialog.");
            panel.Close();
            Stopwatch wait = Stopwatch.StartNew();
            while (!Enabled && wait.ElapsedMilliseconds < 3000) await Task.Delay(20);
            if (!Enabled) throw new InvalidOperationException("Main window remained disabled after closing the panel.");
        }
        void DemandVirtual()
        {
            if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Count == 0 || _blocks.Any(b => !b.Device.IsTestDevice))
                throw new InvalidOperationException("The hardware matrix lost its virtual-device boundary.");
        }
        await OpenPanel();
        string[] systems = _devMatrixSystems.ToArray(), cpus = _devMatrixCpus.ToArray();
        if (systems.Length == 0 || cpus.Length == 0) throw new InvalidOperationException("Preset catalogs are empty.");
        var profiles = systems.Select((name, index) => (Kind: "system", Name: name, Index: index))
            .Concat(cpus.Select((name, index) => (Kind: "cpu", Name: name, Index: index)))
            .Where(profile => !cpuOnly || profile.Kind == "cpu").ToArray();
        HashSet<string> selected = ProfileSelection.Parse(profileFilter,
            profiles.Select(profile => profile.Kind + "-" + profile.Index.ToString("D2")));
        profiles = profiles.Where(profile => selected.Contains(profile.Kind + "-" + profile.Index.ToString("D2"))).ToArray();
        List<object> totals = [];
        int totalPassed = 0, totalFailed = 0, totalNotApplicable = 0, profileFailures = 0;
        foreach (var profile in profiles)
        {
            if (_devMatrixLoadSystem is null) await OpenPanel();
            string key = profile.Kind + "-" + profile.Index.ToString("D2");
            Stopwatch clock = Stopwatch.StartNew();
            List<string> failures = [];
            List<object> cases = [];
            int passed = 0, failed = 0, notApplicable = 0;
            WriteLog($"TEST.HARDWARE.PROFILE.BEGIN: id={key} name=\"{profile.Name}\"");
            if (!_devMatrixLoadSystem!(profile.Kind == "system" ? profile.Name : systems[0]))
                throw new InvalidOperationException("System preset loader rejected " + profile.Name);
            await _lastDeviceRefresh;
            DemandVirtual();
            // Clear AUTO preview edits before treating each profile as a fresh machine.
            foreach (DeviceBlock block in _blocks) LoadBlockSettings(block);
            TestCpuConfig? expectedCpu = null;
            if (profile.Kind == "cpu")
            {
                expectedCpu = _devMatrixLoadCpu!(profile.Name);
                await _lastDeviceRefresh;
            }
            OperationReport initial = await RefreshBlocksAsync(preserveEdits: false);
            if (initial.Errors.Count != 0) failures.Add("Initial model refresh failed.");
            DemandVirtual();
            int logical = _cpuInfo?.Topology.Logical ?? 0;
            if (logical is < 1 or > 64) failures.Add("Invalid modeled logical processor count: " + logical);
            if (expectedCpu is not null)
            {
                bool topologyMatches = _testCpuActive && logical == expectedCpu.LogicalCount
                    && _testCpuName == expectedCpu.CpuName.Trim()
                    && _cpuInfo!.Topology.LPs.All(lp => expectedCpu.CoreMap[lp.LP] == lp.Core
                        && expectedCpu.ECoreLps.Contains(lp.LP) == (lp.EffClass == 0))
                    && expectedCpu.CcdMap!.All(pair => _cpuInfo.CcdMap[pair.Key] == pair.Value)
                    && expectedCpu.CcxMap!.All(pair => _cpuInfo.CcxMap[pair.Key] == pair.Value);
                if (!topologyMatches) throw new InvalidOperationException("CPU preset was not applied to the modeled topology: " + profile.Name);
            }
            Dictionary<string, TestDeviceState> baseline = CaptureTestSandboxState();
            bool imod = _blocks.Any(b => IsUsbImodTarget(b.Device));
            // Use the existing applicability rules and shared scenario runner.
            foreach (TestSandboxOperation operation in presentationOnly ? Array.Empty<TestSandboxOperation>() : Enum.GetValues<TestSandboxOperation>())
            foreach (TestSandboxFault fault in Enum.GetValues<TestSandboxFault>())
            {
                RestoreTestSandboxState(baseline);
                _testSandbox.Backup = null;
                foreach (DeviceBlock block in _blocks) LoadBlockSettings(block);
                TestSandboxScenario scenario = CaptureCurrentTestScenario(key + "/" + operation + "/" + fault, operation, fault, imod);
                if (!IsTestFaultApplicable(scenario))
                {
                    notApplicable++;
                    cases.Add(new { operation = operation.ToString(), fault = fault.ToString(), status = "NOT_APPLICABLE" });
                    continue;
                }
                TestSandboxRunResult result = RunTestSandboxScenario(scenario, showResult: false);
                bool ok = result.Passed && result.AssertionsPassed > 0 && result.BlockedRealWrites == 0;
                if (ok) passed++; else { failed++; failures.Add(operation + "/" + fault); }
                cases.Add(new { operation = operation.ToString(), fault = fault.ToString(), status = ok ? "PASS" : "FAIL",
                    result.AssertionsPassed, result.AssertionsFailed, result.VirtualWrites, result.BlockedRealWrites,
                    expectedOperationMessages = result.Report.Errors.ToArray() });
                await _lastDeviceRefresh;
                await Task.Delay(1);
            }
            RestoreTestSandboxState(baseline);
            _testSandbox.Backup = null;
            foreach (DeviceBlock block in _blocks) LoadBlockSettings(block);
            await ClosePanel();
            UiLanguage.Set(profile.Index % 2 == 0 ? UiLanguageCode.English : UiLanguageCode.Russian, persist: false);
            OperationReport refresh = await RefreshBlocksAsync(preserveEdits: false);
            if (refresh.Errors.Count != 0) failures.Add("REFRESH after the operation matrix failed.");
            DemandVirtual();
            DeviceBlock editable = _blocks.First(b => !b.Device.Wifi && b.Kind != DeviceKind.STOR);
            editable.LimitBox.Text = "2";
            var drafts = _blocks.ToDictionary(b => b.Device.InstanceId, CaptureEditorState);
            refresh = await RefreshBlocksAsync();
            if (refresh.Errors.Count != 0) failures.Add("REFRESH with a pending edit failed.");
            foreach (DeviceBlock block in _blocks)
            {
                if (CaptureEditorState(block) != drafts[block.Device.InstanceId])
                    failures.Add("Pending editor state changed: " + block.Device.InstanceId);
                if (logical < 64 && (block.AffinityMask >> logical) != 0)
                    failures.Add("Affinity mask exceeds the modeled CPU count: " + block.Device.InstanceId);
            }
            // Validate the completed animated navigation on the enabled main window.
            var visible = _blocks.ToDictionary(b => b.Device.InstanceId, b => b.Group.Visible);
            string category = _filterCategoryButtons.Select(b => b.Tag as string ?? "ALL")
                .Where(c => c != "ALL")
                .OrderByDescending(c => _blocks.FirstOrDefault(b => b.Group.Visible && MatchesCategory(b, c))?.Group.Top ?? -1)
                .First();
            DeviceBlock target = _blocks.First(b => b.Group.Visible && MatchesCategory(b, category));
            int expectedOffset = Math.Min(Math.Max(0, target.Group.Top - UiScale(12)),
                Math.Max(0, _devicesScroll.Maximum - _devicesScroll.ViewportSize));
            SetCategoryFilter(category);
            Stopwatch navigation = Stopwatch.StartNew();
            while (_devicesScroll.IsScrolling && navigation.ElapsedMilliseconds < 1000) await Task.Delay(16);
            int actualOffset = _devicesScroll.Value;
            bool navigationPassed = !_devicesScroll.IsScrolling && actualOffset == expectedOffset
                && _activeCategoryFilter == category && _blocks.All(b => b.Group.Visible == visible[b.Device.InstanceId]);
            if (!navigationPassed) failures.Add($"Completed navigation failed: {category}, expected={expectedOffset}, actual={actualOffset}");
            SetCategoryFilter("ALL");
            navigation.Restart();
            while (_devicesScroll.IsScrolling && navigation.ElapsedMilliseconds < 1000) await Task.Delay(16);
            if (_devicesScroll.IsScrolling || _devicesScroll.Value != 0) failures.Add("Navigation did not return to ALL/top.");
            DevCaptureWindow(key, "@main", false);
            string status = failed == 0 && failures.Count == 0 ? "PASS" : "FAIL";
            if (status == "FAIL") profileFailures++;
            string relative = "state/" + key + ".json";
            DevFiles.Write(DevArtifact(relative), new { status, profile.Kind, profile.Name, logicalProcessors = logical,
                modelCpu = new { name = _testCpuName, logical, topology = _cpuInfo!.Topology.LPs,
                    _cpuInfo.CcdMap, _cpuInfo.CcxMap, _cpuInfo.CcdL3CacheBytes },
                devices = _blocks.Select(b => new { b.Device.Name, kind = b.Kind.ToString(), b.Device.Wifi, b.Device.IsTestDevice }).ToArray(),
                passed, failed, notApplicable, failures, cases, presentationOnly,
                navigation = new { category, expectedOffset, actualOffset, passed = navigationPassed }, elapsedMs = clock.Elapsed.TotalMilliseconds,
                scope = "Existing test presets and shared operation runner; virtual writes only; no hardware driver loading" });
            DevRecordState(relative);
            totals.Add(new { id = key, profile.Kind, profile.Name, status, logical, devices = _blocks.Count,
                passed, failed, notApplicable, failures, report = relative });
            totalPassed += passed; totalFailed += failed; totalNotApplicable += notApplicable;
            WriteLog($"TEST.HARDWARE.PROFILE.END: id={key} status={status} passed={passed} failed={failed} notApplicable={notApplicable}");
        }
        string summary = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(summary), new { status = profileFailures == 0 ? "PASS" : "FAIL",
            systemProfiles = systems.Length, cpuProfiles = cpus.Length, selectedProfiles = profiles.Length, totalPassed, totalFailed, totalNotApplicable, presentationOnly, cpuOnly, profileFilter,
            profileFailures, profiles = totals, hardwareValidation = "MODELED_ONLY" });
        DevRecordState(summary);
        if (profileFailures != 0)
            throw new InvalidOperationException("Hardware matrix found failing profiles: " + profileFailures + ". See state/" + id + ".json");
    }
}
