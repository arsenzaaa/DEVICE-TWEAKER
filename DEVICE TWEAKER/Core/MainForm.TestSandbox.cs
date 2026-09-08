using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private TestDeviceState EnsureTestDeviceState(DeviceInfo device)
    {
        if (device.TestState is not null)
        {
            return device.TestState;
        }

        bool powerSaving = device.Kind == DeviceKind.USB
            ? !string.Equals(device.UsbSelectiveSuspend, "off", StringComparison.OrdinalIgnoreCase)
            : !string.Equals(device.NicPowerSaving, "off", StringComparison.OrdinalIgnoreCase);
        device.TestState = new TestDeviceState
        {
            MsiEnabled = !string.Equals(device.TestMsiStatus, "Disabled", StringComparison.OrdinalIgnoreCase),
            PowerSavingEnabled = device.Wifi ? null : powerSaving,
        };
        return device.TestState;
    }

    private bool TryConsumeTestFault(TestSandboxFault fault)
    {
        if (!_testSandbox.Active || _testSandbox.FaultConsumed || _testSandbox.Fault != fault)
        {
            return false;
        }

        _testSandbox.FaultConsumed = true;
        WriteLog($"TEST.SANDBOX.FAULT: injected={fault}");
        return true;
    }

    private bool ApplyTestBlockSettings(DeviceBlock block, bool autoMsiOnly, OperationReport? report)
    {
        TestDeviceState state = EnsureTestDeviceState(block.Device);
        if (TryConsumeTestFault(TestSandboxFault.RegistryWriteDenied))
        {
            const string technical = "UnauthorizedAccessException: simulated HKLM write access denied.";
            report?.AddError($"{block.Device.Name} — registry", "Windows denied access to the device settings.", technical);
            WriteLog($"TEST.SANDBOX.WRITE.ERROR: id={block.Device.InstanceId} boundary=registry error=access-denied");
            return false;
        }

        state.MsiEnabled = string.Equals(block.MsiCombo.SelectedItem?.ToString(), "Enabled", StringComparison.OrdinalIgnoreCase);
        _testSandbox.VirtualWrites++;
        if (autoMsiOnly)
        {
            WriteLog($"TEST.SANDBOX.WRITE: id={block.Device.InstanceId} MSI={state.MsiEnabled} mode=msi-only");
            return true;
        }

        if (!TryParseMsiLimitInput(block.LimitBox.Text, out bool unlocked, out int limit))
        {
            report?.AddError($"{block.Device.Name} — MSI Limit", "invalid value; expected 0 or a whole number from 1 to 2048");
            return false;
        }

        state.MsiLimit = unlocked ? null : limit;
        state.Priority = block.PrioCombo.SelectedItem?.ToString() switch
        {
            "Low" => 1,
            "Normal" => 2,
            "High" => 3,
            _ => null,
        };

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            if (TryConsumeTestFault(TestSandboxFault.NicItrReadFailed)
                || TryConsumeTestFault(TestSandboxFault.NicItrWriteFailed))
            {
                report?.AddError(
                    $"{block.Device.Name} — NIC ITR",
                    "The NIC interrupt moderation value could not be verified.",
                    _testSandbox.Fault == TestSandboxFault.NicItrReadFailed
                        ? "DeviceIoControl(NIC_ITR_READ) failed with simulated Win32 error 31."
                        : "DeviceIoControl(NIC_ITR_WRITE) failed with simulated Win32 error 5.");
                return false;
            }

            state.RssBaseCore = block.RssBaseCore;
            state.RssQueues = Math.Max(1, (int)(block.RssQueueBox?.Value ?? 1));
            state.NdisMode = block.NdisModeCombo?.SelectedItem?.ToString() ?? "RSS";
            state.AffinityMask = block.AffinityMask;
            state.Policy = state.NdisMode is "IRQ" or "BOTH" ? 4 : 0;
        }
        else if (block.Kind == DeviceKind.STOR)
        {
            state.Policy = 0;
            state.AffinityMask = 0;
        }
        else
        {
            state.Policy = MapPolicyText(block.PolicyCombo.SelectedItem?.ToString() ?? "MachineDefault") ?? 0;
            state.AffinityMask = state.Policy == 0 ? 0 : block.AffinityMask;
        }

        if (block.PowerSavingCheck is not null && !block.Device.Wifi)
        {
            if (TryConsumeTestFault(TestSandboxFault.UsbPowerWriteFailed)
                || TryConsumeTestFault(TestSandboxFault.WmiTimeout))
            {
                report?.AddError(
                    $"{block.Device.Name} — power saving",
                    "The power setting could not be written.",
                    _testSandbox.Fault == TestSandboxFault.WmiTimeout
                        ? "TimeoutException: simulated MSPower_DeviceEnable WMI timeout."
                        : "Win32Exception: simulated USB power-policy write failure.");
                return false;
            }

            state.PowerSavingEnabled = block.PowerSavingCheck.Checked;
            if (block.Kind == DeviceKind.USB)
            {
                block.Device.UsbSelectiveSuspend = block.PowerSavingCheck.Checked ? "on" : "off";
            }
            else
            {
                block.Device.NicPowerSaving = block.PowerSavingCheck.Checked ? "on" : "off";
            }
        }

        _testSandbox.VirtualWrites += 6;
        WriteLog(
            $"TEST.SANDBOX.WRITE: id={block.Device.InstanceId} MSI={state.MsiEnabled} " +
            $"limit={(state.MsiLimit?.ToString() ?? "absent")} priority={(state.Priority?.ToString() ?? "absent")} " +
            $"policy={state.Policy} mask=0x{state.AffinityMask:X} rssBase={(state.RssBaseCore?.ToString() ?? "absent")} " +
            $"rssQueues={state.RssQueues} power={(state.PowerSavingEnabled?.ToString() ?? "preserved")}");
        return true;
    }

    private void ResetTestBlockState(DeviceBlock block)
    {
        TestDeviceState state = EnsureTestDeviceState(block.Device);
        state.Priority = null;
        state.Policy = 0;
        state.AffinityMask = 0;
        state.RssBaseCore = null;
        state.RssQueues = 1;
        state.NdisMode = "RSS";
        if (!block.Device.Wifi)
        {
            state.PowerSavingEnabled = true;
        }

        state.ImodValue = "0x0";
        state.NicItrValue = "default";
        _testSandbox.VirtualWrites += 8;
        WriteLog($"TEST.SANDBOX.RESET: id={block.Device.InstanceId} MSI/limit=preserved other-managed-state=default");
    }

    private Dictionary<string, TestDeviceState> CaptureTestSandboxState()
    {
        return _testDevices.ToDictionary(
            device => device.InstanceId,
            device => EnsureTestDeviceState(device).Clone(),
            StringComparer.OrdinalIgnoreCase);
    }

    private void RestoreTestSandboxState(IReadOnlyDictionary<string, TestDeviceState> snapshot)
    {
        foreach (DeviceInfo device in _testDevices)
        {
            if (snapshot.TryGetValue(device.InstanceId, out TestDeviceState? state))
            {
                device.TestState = state.Clone();
            }
        }
    }

    private void AddSimulatedImodResult(OperationReport report)
    {
        TestSandboxFault fault = _testSandbox.Fault;
        string? technical = fault switch
        {
            TestSandboxFault.KduMissing => "kdu.exe is missing from the verified embedded payload.",
            TestSandboxFault.KduTimeout => "kdu.exe timed out after 60000 ms and was terminated.",
            TestSandboxFault.KduExitCode => "kdu.exe exited with code 1. stdout: simulated mapper failure. stderr: vulnerable driver was blocked.",
            TestSandboxFault.KduDeviceUnavailable577 => "kdu.exe exited successfully but \\\\.\\DeviceTweakerImod2 is unavailable. Service fallback failed: Windows cannot verify the digital signature for this file. (code 577) (kernel CI blocked; enabled=True testSign=False testBuild=False hvciRuntime=False)",
            TestSandboxFault.ImodReadFailed => "DeviceIoControl(IMOD_READ) failed with Win32 error 31.",
            TestSandboxFault.ImodWriteFailed => "DeviceIoControl(IMOD_WRITE) failed with Win32 error 5.",
            TestSandboxFault.ImodVerificationMismatch => "IMOD verification mismatch: requested=0x8 actual=0x0.",
            _ => null,
        };

        if (technical is null)
        {
            foreach (DeviceInfo device in _testDevices.Where(IsUsbImodTarget))
            {
                DeviceBlock? block = _blocks.FirstOrDefault(candidate =>
                    string.Equals(candidate.Device.InstanceId, device.InstanceId, StringComparison.OrdinalIgnoreCase));
                EnsureTestDeviceState(device).ImodValue = string.IsNullOrWhiteSpace(block?.ImodBox.Text)
                    ? "0x0"
                    : block.ImodBox.Text.Trim();
                _testSandbox.VirtualWrites++;
            }
            report.AddSuccess("USB IMOD", "Applied");
            return;
        }

        _testSandbox.FaultConsumed = true;
        WriteLog($"TEST.SANDBOX.IMOD.ERROR: fault={fault} detail=\"{FlattenLogText(technical)}\"");
        report.AddError(
            "USB IMOD",
            FormatImodUnavailableUserMessage(technical, includeNotChanged: true),
            technical);
    }

    private bool IsTestFaultApplicable(TestSandboxScenario scenario)
    {
        bool apply = scenario.Operation is TestSandboxOperation.Auto or TestSandboxOperation.Apply;
        return scenario.Fault switch
        {
            TestSandboxFault.None => true,
            TestSandboxFault.CpuTopologyUnavailable or TestSandboxFault.CpuMapUnavailable => scenario.Operation == TestSandboxOperation.Auto,
            TestSandboxFault.BackupAccessDenied => scenario.Operation is TestSandboxOperation.Auto or TestSandboxOperation.Apply or TestSandboxOperation.Backup,
            TestSandboxFault.RegistryWriteDenied => apply,
            TestSandboxFault.UsbPowerWriteFailed or TestSandboxFault.WmiTimeout => apply && _blocks.Any(block => block.PowerSavingCheck is not null && !block.Device.Wifi),
            TestSandboxFault.KduMissing or TestSandboxFault.KduTimeout or TestSandboxFault.KduExitCode
                or TestSandboxFault.KduDeviceUnavailable577 or TestSandboxFault.ImodReadFailed
                or TestSandboxFault.ImodWriteFailed or TestSandboxFault.ImodVerificationMismatch =>
                    apply && scenario.OptimizeUsbImod && _blocks.Any(block => IsUsbImodTarget(block.Device)),
            TestSandboxFault.NicItrReadFailed or TestSandboxFault.NicItrWriteFailed =>
                apply && _blocks.Any(block => block.Kind == DeviceKind.NET_NDIS && !block.Device.Wifi),
            TestSandboxFault.CorruptBackup or TestSandboxFault.RestoreWriteFailed or TestSandboxFault.RollbackFailed =>
                scenario.Operation == TestSandboxOperation.Restore,
            TestSandboxFault.RefreshFailed => apply,
            _ => false,
        };
    }

    private TestSandboxRunResult RunTestSandboxScenario(TestSandboxScenario scenario, bool showResult)
    {
        OperationReport report = new();
        TestSandboxRunResult result = new()
        {
            Report = report,
            ScenarioName = scenario.Name,
            Operation = scenario.Operation,
            Fault = scenario.Fault,
        };

        bool previousDryRun = _testAutoDryRun;
        bool previousOnly = _testDevicesOnly;
        bool previousTopologyReliable = _cpuTopologyReliable;
        CpuInfo? previousCpuInfo = _cpuInfo;
        _testAutoDryRun = true;
        _testDevicesOnly = true;
        _testSandbox.Begin(scenario.Fault);
        WriteLog($"TEST.SCENARIO.BEGIN: name=\"{FlattenLogText(scenario.Name)}\" operation={scenario.Operation} fault={scenario.Fault} devices={_testDevices.Count}");

        try
        {
            if (_blocks.Count == 0 || _blocks.Any(block => !block.Device.IsTestDevice))
            {
                RefreshBlocks();
            }

            bool allVirtual = _blocks.Count > 0 && _blocks.All(block => block.Device.IsTestDevice);
            if (allVirtual)
            {
                result.AssertionsPassed++;
            }
            else
            {
                result.AssertionsFailed++;
                report.AddError("SANDBOX", "The test scenario contains a real device; execution was stopped.");
                report.MarkNoChangesMade();
                goto Finished;
            }

            Dictionary<string, TestDeviceState> before = CaptureTestSandboxState();
            foreach ((string id, TestDeviceState state) in before)
            {
                WriteLog($"TEST.SCENARIO.STATE.INITIAL: id={id} {FormatTestStateForLog(state)}");
            }
            bool faultApplicable = IsTestFaultApplicable(scenario);
            if (!faultApplicable)
            {
                result.AssertionsFailed++;
                report.MarkNoChangesMade();
                report.AddError(
                    "TEST SCENARIO",
                    $"Failure {scenario.Fault} does not apply to {scenario.Operation} or the selected virtual hardware.");
                goto Finished;
            }
            if (scenario.Operation is TestSandboxOperation.Auto or TestSandboxOperation.Apply)
            {
                if (TryConsumeTestFault(TestSandboxFault.BackupAccessDenied))
                {
                    report.MarkNoChangesMade();
                    report.AddError("AUTOMATIC BACKUP", "The backup could not be created. No changes were made.", "UnauthorizedAccessException: simulated access denied while creating the pre-operation backup.");
                    goto Verify;
                }

                _testSandbox.Backup = CaptureTestSandboxState();
                report.AddSuccess("BACKUP", "Virtual snapshot created");
            }

            if (scenario.Operation == TestSandboxOperation.Auto)
            {
                if (TryConsumeTestFault(TestSandboxFault.CpuTopologyUnavailable))
                {
                    _cpuTopologyReliable = false;
                }
                else if (TryConsumeTestFault(TestSandboxFault.CpuMapUnavailable))
                {
                    _cpuInfo = null;
                }
                bool hasImod = _blocks.Any(block => IsUsbImodTarget(block.Device));
                if (!InvokeAutoOptimization(scenario.OptimizeUsbImod, hasImod, report))
                {
                    report.MarkNoChangesMade();
                    goto Verify;
                }
                report.AddSuccess("AUTO PLAN", "Built");
            }

            if (scenario.Operation is TestSandboxOperation.Auto or TestSandboxOperation.Apply)
            {
                int applied = 0;
                foreach (DeviceBlock block in _blocks)
                {
                    if (block.Device.Wifi)
                    {
                        WriteLog($"TEST.SANDBOX.PRESERVE: id={block.Device.InstanceId} reason=wifi");
                        continue;
                    }

                    if (ApplyTestBlockSettings(block, scenario.Operation == TestSandboxOperation.Auto && IsAutoMsiOnlyDevice(block), report))
                    {
                        applied++;
                    }
                }
                report.AddSuccess("DEVICE SETTINGS", $"{applied} virtual devices");

                if (TryConsumeTestFault(TestSandboxFault.RefreshFailed))
                {
                    report.AddError("REFRESH", "Settings were applied, but the interface could not be refreshed.", "InvalidOperationException: simulated device refresh failure.");
                }

                if (scenario.OptimizeUsbImod && _blocks.Any(block => IsUsbImodTarget(block.Device)))
                {
                    AddSimulatedImodResult(report);
                }
            }
            else if (scenario.Operation == TestSandboxOperation.SafeReset)
            {
                foreach (DeviceBlock block in _blocks.Where(block => !block.Device.Wifi))
                {
                    ResetBlockSettings(block, report);
                }
                report.AddSuccess("RESET WINDOWS DEFAULT", "Virtual state reset");
            }
            else if (scenario.Operation == TestSandboxOperation.Backup)
            {
                if (TryConsumeTestFault(TestSandboxFault.BackupAccessDenied))
                {
                    report.MarkNoChangesMade();
                    report.AddError("BACKUP", "The virtual backup could not be created.", "UnauthorizedAccessException: simulated backup directory access denied.");
                }
                else
                {
                    _testSandbox.Backup = CaptureTestSandboxState();
                    report.AddSuccess("BACKUP", $"{_testSandbox.Backup.Count} devices captured");
                }
            }
            else if (scenario.Operation == TestSandboxOperation.Restore)
            {
                _testSandbox.Backup ??= before.ToDictionary(pair => pair.Key, pair => pair.Value.Clone(), StringComparer.OrdinalIgnoreCase);
                if (TryConsumeTestFault(TestSandboxFault.CorruptBackup))
                {
                    report.MarkNoChangesMade();
                    report.AddError("BACKUP RESTORE", "The selected backup is invalid.", "JsonException: simulated corrupt JSON backup.");
                }
                else if (TryConsumeTestFault(TestSandboxFault.RestoreWriteFailed))
                {
                    report.AddError("BACKUP RESTORE", "Restore stopped while writing device settings. Rollback completed.", "UnauthorizedAccessException: simulated registry failure during restore.");
                }
                else if (TryConsumeTestFault(TestSandboxFault.RollbackFailed))
                {
                    report.AddError("BACKUP RESTORE", "Restore and rollback were incomplete.", "Restore error: simulated write failure. Rollback error: simulated access denied.");
                }
                else
                {
                    RestoreTestSandboxState(_testSandbox.Backup);
                    _testSandbox.VirtualWrites += _testSandbox.Backup.Count;
                    report.AddSuccess("BACKUP RESTORE", $"{_testSandbox.Backup.Count} devices restored");
                }
            }

        Verify:
            bool expectedError = scenario.Fault != TestSandboxFault.None;
            bool outcomeMatches = expectedError ? !report.Succeeded : report.Succeeded;
            if (outcomeMatches) result.AssertionsPassed++; else result.AssertionsFailed++;

            if (scenario.Fault == TestSandboxFault.None || _testSandbox.FaultConsumed) result.AssertionsPassed++; else result.AssertionsFailed++;

            bool wifiPreserved = _testDevices.Where(device => device.Wifi).All(device =>
                before.TryGetValue(device.InstanceId, out TestDeviceState? initial)
                && StatesEqual(initial, EnsureTestDeviceState(device)));
            if (wifiPreserved) result.AssertionsPassed++; else result.AssertionsFailed++;

            if (report.NoChangesMade)
            {
                bool unchanged = _testDevices.All(device =>
                    before.TryGetValue(device.InstanceId, out TestDeviceState? initial)
                    && StatesEqual(initial, EnsureTestDeviceState(device)));
                if (unchanged) result.AssertionsPassed++; else result.AssertionsFailed++;
            }

            if (scenario.Fault == TestSandboxFault.None
                && scenario.Operation is TestSandboxOperation.Auto or TestSandboxOperation.Apply)
            {
                bool finalStateMatchesUi = _blocks.Where(block => !block.Device.Wifi).All(block =>
                {
                    TestDeviceState state = EnsureTestDeviceState(block.Device);
                    bool settingsMatch = TestStateMatchesBlock(
                        state, block, scenario.Operation == TestSandboxOperation.Auto && IsAutoMsiOnlyDevice(block));
                    bool imodMatches = !scenario.OptimizeUsbImod
                        || !IsUsbImodTarget(block.Device)
                        || string.Equals(state.ImodValue, block.ImodBox.Text?.Trim(), StringComparison.Ordinal);
                    return settingsMatch && imodMatches;
                });
                if (finalStateMatchesUi) result.AssertionsPassed++; else result.AssertionsFailed++;
            }

            if (scenario.Fault == TestSandboxFault.None && scenario.Operation == TestSandboxOperation.SafeReset)
            {
                bool resetMatches = _testDevices.All(device =>
                {
                    TestDeviceState state = EnsureTestDeviceState(device);
                    if (!before.TryGetValue(device.InstanceId, out TestDeviceState? initial)) return false;
                    if (device.Wifi) return StatesEqual(initial, state);
                    return state.MsiEnabled == initial.MsiEnabled
                        && state.MsiLimit == initial.MsiLimit
                        && state.Priority is null
                        && state.Policy == 0
                        && state.AffinityMask == 0
                        && state.RssBaseCore is null
                        && state.RssQueues == 1
                        && state.PowerSavingEnabled != false;
                });
                if (resetMatches) result.AssertionsPassed++; else result.AssertionsFailed++;
            }

            if (scenario.Fault == TestSandboxFault.None && scenario.Operation == TestSandboxOperation.Backup)
            {
                bool backupMatches = _testSandbox.Backup is not null
                    && before.Count == _testSandbox.Backup.Count
                    && before.All(pair => _testSandbox.Backup.TryGetValue(pair.Key, out TestDeviceState? saved) && StatesEqual(pair.Value, saved));
                if (backupMatches) result.AssertionsPassed++; else result.AssertionsFailed++;
            }

            if (scenario.Fault == TestSandboxFault.None && scenario.Operation == TestSandboxOperation.Restore)
            {
                bool restoreMatches = _testSandbox.Backup is not null
                    && _testDevices.All(device => _testSandbox.Backup.TryGetValue(device.InstanceId, out TestDeviceState? saved)
                        && StatesEqual(saved, EnsureTestDeviceState(device)));
                if (restoreMatches) result.AssertionsPassed++; else result.AssertionsFailed++;
            }

            if (_testAutoDryRun && _testDevicesOnly) result.AssertionsPassed++; else result.AssertionsFailed++;
            if (_testSandbox.BlockedRealWrites == 0) result.AssertionsPassed++; else result.AssertionsFailed++;

            if (result.AssertionsFailed == 0)
            {
                report.AddSuccess("SANDBOX ASSERTIONS", $"{result.AssertionsPassed} passed");
            }
            else
            {
                report.AddError("SANDBOX ASSERTIONS", $"{result.AssertionsFailed} failed", $"Passed={result.AssertionsPassed}; Failed={result.AssertionsFailed}");
            }

        Finished:
            result.VirtualWrites = _testSandbox.VirtualWrites;
            result.BlockedRealWrites = _testSandbox.BlockedRealWrites;
            foreach (DeviceInfo device in _testDevices)
            {
                WriteLog($"TEST.SCENARIO.STATE.FINAL: id={device.InstanceId} {FormatTestStateForLog(EnsureTestDeviceState(device))}");
            }
            WriteLog(
                $"TEST.SCENARIO.END: name=\"{FlattenLogText(scenario.Name)}\" status={(result.Passed ? "PASS" : "FAIL")} " +
                $"assertions={result.AssertionsPassed}/{result.AssertionsFailed} errors={report.Errors.Count} " +
                $"virtualWrites={result.VirtualWrites} realWrites=0 faultConsumed={_testSandbox.FaultConsumed}");
        }
        catch (Exception ex)
        {
            result.AssertionsFailed++;
            report.AddError("TEST SCENARIO", "The scenario runner failed unexpectedly.", ex.ToString());
            WriteLog($"TEST.SCENARIO.CRASH: {FlattenLogText(ex.ToString())}");
        }
        finally
        {
            _testSandbox.End();
            _testAutoDryRun = previousDryRun;
            _testDevicesOnly = previousOnly;
            _cpuTopologyReliable = previousTopologyReliable;
            _cpuInfo = previousCpuInfo;
        }

        if (showResult)
        {
            ShowOperationResult(
                report,
                $"Scenario passed. Assertions: {result.AssertionsPassed}.",
                $"Scenario completed with errors. Assertions: {result.AssertionsPassed} passed, {result.AssertionsFailed} failed.",
                operationName: $"SANDBOX {scenario.Operation.ToString().ToUpperInvariant()}");
        }

        return result;
    }

    private static bool StatesEqual(TestDeviceState left, TestDeviceState right)
    {
        return left.MsiEnabled == right.MsiEnabled
            && left.MsiLimit == right.MsiLimit
            && left.Priority == right.Priority
            && left.Policy == right.Policy
            && left.AffinityMask == right.AffinityMask
            && left.RssBaseCore == right.RssBaseCore
            && left.RssQueues == right.RssQueues
            && string.Equals(left.NdisMode, right.NdisMode, StringComparison.Ordinal)
            && left.PowerSavingEnabled == right.PowerSavingEnabled
            && string.Equals(left.ImodValue, right.ImodValue, StringComparison.Ordinal)
            && string.Equals(left.NicItrValue, right.NicItrValue, StringComparison.Ordinal);
    }

    private bool TestStateMatchesBlock(TestDeviceState state, DeviceBlock block, bool msiOnly)
    {
        bool msiMatches = state.MsiEnabled == string.Equals(
            block.MsiCombo.SelectedItem?.ToString(), "Enabled", StringComparison.OrdinalIgnoreCase);
        if (!msiMatches || msiOnly)
        {
            return msiMatches;
        }

        bool limitMatches = TryParseMsiLimitInput(block.LimitBox.Text, out bool unlocked, out int limit)
            && state.MsiLimit == (unlocked ? null : limit);
        int? priority = block.PrioCombo.SelectedItem?.ToString() switch
        {
            "Low" => 1,
            "Normal" => 2,
            "High" => 3,
            _ => null,
        };
        if (!limitMatches || state.Priority != priority)
        {
            return false;
        }

        if (block.Kind == DeviceKind.NET_NDIS)
        {
            return state.RssBaseCore == block.RssBaseCore
                && state.RssQueues == Math.Max(1, (int)(block.RssQueueBox?.Value ?? 1))
                && string.Equals(state.NdisMode, block.NdisModeCombo?.SelectedItem?.ToString() ?? "RSS", StringComparison.Ordinal);
        }

        if (block.Kind == DeviceKind.STOR)
        {
            return state.Policy == 0 && state.AffinityMask == 0;
        }

        int policy = MapPolicyText(block.PolicyCombo.SelectedItem?.ToString() ?? "MachineDefault") ?? 0;
        return state.Policy == policy
            && state.AffinityMask == (policy == 0 ? 0 : block.AffinityMask)
            && (block.PowerSavingCheck is null || state.PowerSavingEnabled == block.PowerSavingCheck.Checked);
    }

    private static string FormatTestStateForLog(TestDeviceState state)
    {
        return $"msi={state.MsiEnabled} limit={(state.MsiLimit?.ToString() ?? "absent")} " +
            $"priority={(state.Priority?.ToString() ?? "absent")} policy={state.Policy} mask=0x{state.AffinityMask:X} " +
            $"rssBase={(state.RssBaseCore?.ToString() ?? "absent")} rssQueues={state.RssQueues} mode={state.NdisMode} " +
            $"power={(state.PowerSavingEnabled?.ToString() ?? "preserve")} imod=\"{state.ImodValue}\" nicItr=\"{state.NicItrValue}\"";
    }

    private TestSandboxScenario CaptureCurrentTestScenario(string name, TestSandboxOperation operation, TestSandboxFault fault, bool optimizeImod)
    {
        TestSandboxScenario scenario = new()
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Untitled scenario" : name.Trim(),
            Operation = operation,
            Fault = fault,
            OptimizeUsbImod = optimizeImod,
            Cpu = new TestScenarioCpu
            {
                Name = _testCpuName,
                LogicalCount = _cpuInfo?.Topology.Logical ?? _maxLogical,
                SmtEnabled = _cpuInfo?.Topology.ByCore.Values.Any(group => group.Count > 1) == true,
                UseHyperThreadingLabel = _smtText.Contains("Hyper-Threading", StringComparison.OrdinalIgnoreCase),
                ECoreLps = _cpuInfo?.Topology.LPs.Where(IsEfficiencyCore).Select(lp => lp.LP).ToList() ?? [],
                CoreMap = _cpuInfo?.Topology.LPs.ToDictionary(lp => lp.LP, lp => lp.Core) ?? [],
                CcdMap = _cpuInfo?.CcdMap.ToDictionary(pair => pair.Key, pair => pair.Value) ?? [],
                CcxMap = _cpuInfo?.CcxMap.ToDictionary(pair => pair.Key, pair => pair.Value) ?? [],
                CppcRatings = _cppcRatings.ToDictionary(pair => pair.Key, pair => pair.Value),
            },
        };

        scenario.Devices = _testDevices.Select(device => new TestScenarioDevice
        {
            Kind = device.Kind,
            Name = device.Name,
            InstanceId = device.InstanceId,
            UsbRoles = device.UsbRoles,
            AudioEndpoints = device.AudioEndpoints,
            StorageTag = device.StorageTag,
            Wifi = device.Wifi,
            UsbIsXhci = device.UsbIsXhci,
            UsbHasDevices = device.UsbHasDevices,
            IntegratedGpu = device.IsIntegratedGpu,
            IrqCount = device.TestIrqCount,
            MsiStatus = device.TestMsiStatus,
            UsbSelectiveSuspend = device.UsbSelectiveSuspend,
            NicPowerSaving = device.NicPowerSaving,
            InitialState = EnsureTestDeviceState(device).Clone(),
        }).ToList();
        return scenario;
    }

    private void LoadTestSandboxScenario(TestSandboxScenario scenario)
    {
        if (scenario.Version != 1 || scenario.Cpu.LogicalCount < 1 || scenario.Cpu.LogicalCount > MaxAffinityBits)
        {
            throw new InvalidOperationException("Unsupported or invalid sandbox scenario.");
        }
        if (scenario.Devices.Count is < 1 or > 256)
        {
            throw new InvalidOperationException("A scenario must contain between 1 and 256 virtual devices.");
        }
        if (scenario.Devices.Any(device => string.IsNullOrWhiteSpace(device.InstanceId) || string.IsNullOrWhiteSpace(device.Name)))
        {
            throw new InvalidOperationException("Every virtual device must have a name and PNP ID.");
        }
        if (scenario.Devices.Select(device => NormalizeInstanceId(device.InstanceId)).Distinct(StringComparer.OrdinalIgnoreCase).Count() != scenario.Devices.Count)
        {
            throw new InvalidOperationException("The scenario contains duplicate PNP IDs.");
        }
        if (scenario.Cpu.CoreMap.Keys.Any(lp => lp < 0 || lp >= scenario.Cpu.LogicalCount)
            || scenario.Cpu.CcdMap.Keys.Any(lp => lp < 0 || lp >= scenario.Cpu.LogicalCount)
            || scenario.Cpu.CcxMap.Keys.Any(lp => lp < 0 || lp >= scenario.Cpu.LogicalCount)
            || scenario.Cpu.ECoreLps.Any(lp => lp < 0 || lp >= scenario.Cpu.LogicalCount))
        {
            throw new InvalidOperationException("The CPU map contains a logical processor outside the configured range.");
        }
        foreach (TestScenarioDevice device in scenario.Devices)
        {
            TestDeviceState state = device.InitialState ?? throw new InvalidOperationException($"Initial state is missing for {device.Name}.");
            if (state.MsiLimit is < 1 or > 2048 || state.RssQueues < 1 || state.RssQueues > MaxAffinityBits)
            {
                throw new InvalidOperationException($"Initial MSI/RSS values are invalid for {device.Name}.");
            }
        }

        TestCpuConfig cpu = new()
        {
            CpuName = scenario.Cpu.Name,
            LogicalCount = scenario.Cpu.LogicalCount,
            SmtEnabled = scenario.Cpu.SmtEnabled,
            UseHyperThreadingLabel = scenario.Cpu.UseHyperThreadingLabel,
            CcdMap = scenario.Cpu.CcdMap,
            CcxMap = scenario.Cpu.CcxMap,
        };
        foreach (int lp in scenario.Cpu.ECoreLps) cpu.ECoreLps.Add(lp);
        foreach ((int lp, int core) in scenario.Cpu.CoreMap) cpu.CoreMap[lp] = core;
        foreach ((int lp, int rating) in scenario.Cpu.CppcRatings) cpu.CppcRatings[lp] = rating;
        ApplyTestCpuConfig(cpu);

        _testDevices.Clear();
        foreach (TestScenarioDevice saved in scenario.Devices)
        {
            DeviceInfo device = CreateTestDevice(
                saved.Kind, saved.Name, saved.InstanceId, saved.UsbRoles, saved.AudioEndpoints, saved.StorageTag,
                saved.Wifi, saved.UsbIsXhci, saved.UsbHasDevices, saved.IntegratedGpu, saved.IrqCount,
                saved.MsiStatus, saved.UsbSelectiveSuspend, saved.NicPowerSaving);
            device.TestState = saved.InitialState.Clone();
            _testDevices.Add(device);
        }

        _testDevicesEnabled = true;
        _testDevicesOnly = true;
        _testAutoDryRun = true;
        _testSandbox.Backup = null;
        RefreshBlocks();
        WriteLog($"TEST.SCENARIO.LOAD: name=\"{FlattenLogText(scenario.Name)}\" devices={scenario.Devices.Count}");
    }

    private static JsonSerializerOptions TestScenarioJsonOptions() => new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };
}
