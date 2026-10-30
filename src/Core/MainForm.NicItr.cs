using System.Globalization;
using System.Runtime.InteropServices;
using System.Security;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    internal static NicItrProfile? TryGetNicItrProfile(string instanceId)
        => NicItrService.TryGetNicItrProfile(instanceId);

    private List<NicDriverModeration.Keyword> ReadNicDriverModeration(string instanceId)
    {
        List<NicDriverModeration.Keyword> keywords = [];
        string? classKey = GetClassKeyForDevice(instanceId);
        if (string.IsNullOrWhiteSpace(classKey))
        {
            return keywords;
        }

        try
        {
            using RegistryKey? adapterKey = Registry.LocalMachine.OpenSubKey(classKey);
            if (adapterKey is null)
            {
                return keywords;
            }

            foreach (string name in NicDriverModeration.KeywordNames)
            {
                using RegistryKey? paramKey = adapterKey.OpenSubKey($@"Ndi\params\{name}");
                if (paramKey is null)
                {
                    continue;
                }

                Dictionary<string, string> options = new(StringComparer.Ordinal);
                using (RegistryKey? enumKey = paramKey.OpenSubKey("enum"))
                {
                    foreach (string value in enumKey?.GetValueNames() ?? [])
                    {
                        options[value] = enumKey!.GetValue(value)?.ToString() ?? value;
                    }
                }

                keywords.Add(new NicDriverModeration.Keyword(
                    name,
                    options,
                    paramKey.GetValue("default")?.ToString(),
                    adapterKey.GetValue(name)?.ToString()));
            }
        }
        catch (Exception ex) when (ex is SecurityException or IOException or UnauthorizedAccessException)
        {
            WriteLog($"NIC.MODERATION.READ.WARN: {instanceId} {ex.Message}");
        }

        return keywords;
    }

    private bool WriteNicDriverModeration(
        string instanceId,
        IReadOnlyList<NicDriverModeration.Change> changes,
        string reason,
        OperationReport? report = null)
    {
        DemandRealWrite("WriteNicDriverModeration");
        if (changes.Count == 0)
        {
            return true;
        }

        string? classKey = GetClassKeyForDevice(instanceId);
        try
        {
            if (string.IsNullOrWhiteSpace(classKey))
            {
                throw new InvalidOperationException("NDIS class key was not found");
            }

            using RegistryKey adapterKey = Registry.LocalMachine.OpenSubKey(classKey, writable: true)
                ?? throw new InvalidOperationException("NDIS class key is unavailable");
            foreach (NicDriverModeration.Change change in changes)
            {
                // Advanced properties are REG_SZ; keep another kind if the INF chose one.
                RegistryValueKind kind = adapterKey.GetValue(change.Name) is null
                    ? RegistryValueKind.String
                    : adapterKey.GetValueKind(change.Name);
                object value = kind == RegistryValueKind.DWord
                    ? int.Parse(change.After, NumberStyles.Integer, CultureInfo.InvariantCulture)
                    : change.After;
                adapterKey.SetValue(change.Name, value, kind);
                WriteLog($"NIC.MODERATION.{reason}: {instanceId} {change.Name} {change.Before ?? "(none)"} -> {change.After} (applies after restart)");
            }

            return true;
        }
        catch (Exception ex)
        {
            WriteLog($"NIC.MODERATION.{reason}.ERROR: {instanceId} {FlattenLogText(ex.ToString())}");
            report?.AddError($"{instanceId}: driver interrupt moderation", ex.Message);
            return false;
        }
    }

    // Removing the saved NIC ITR profile hands moderation back to the driver.
    private void RestoreNicDriverModerationForSavedProfiles(OperationReport? report)
    {
        HashSet<string> savedKeys = new(
            ReadSavedStartupConfig(out _).NicItrEntries.Select(entry => entry.Hwid),
            StringComparer.OrdinalIgnoreCase);

        foreach (DeviceBlock block in _blocks.Where(block => block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX
            && !block.Device.IsTestDevice
            && savedKeys.Contains(GetNicItrPersistenceKey(block.Device.InstanceId))))
        {
            List<NicDriverModeration.Change> changes = NicDriverModeration.PlanRestoreDefault(ReadNicDriverModeration(block.Device.InstanceId));
            WriteNicDriverModeration(block.Device.InstanceId, changes, "RESTORE", report);
        }
    }

    private async void RefreshNicItrBlock(DeviceBlock block, bool showReadingStatus = true)
    {
        if (block.NicItrBox is null || block.NicItrStatusLabel is null)
        {
            return;
        }

        NicItrProfile? profile = TryGetNicItrProfile(block.Device.InstanceId);
        if (profile is null)
        {
            block.NicItrBox.Text = string.Empty;
            block.NicItrStatusLabel.Text = "current: unsupported";
            if (block.NicItrTimeLabel is not null)
            {
                block.NicItrTimeLabel.Text = "time: unsupported";
                block.NicItrTimeLabel.ForeColor = _statusInactive;
            }
            block.NicItrStatusLabel.ForeColor = _statusInactive;
            SetNicItrTooltip(block, "NIC ITR is unsupported for this adapter.");
            return;
        }

        if (block.Device.IsTestDevice)
        {
            TestDeviceState testState = EnsureTestDeviceState(block.Device);
            List<ulong> previewValues;
            if (string.IsNullOrWhiteSpace(testState.NicItrValue)
                || string.Equals(testState.NicItrValue, "default", StringComparison.OrdinalIgnoreCase)
                || !TryParseNicItrInput(testState.NicItrValue, profile, out previewValues))
            {
                previewValues = Enumerable.Repeat(0UL, Math.Max(1, profile.MaxQueues)).ToList();
            }
            string previewText = FormatNicItrValueList(previewValues, profile);
            block.NicItrBox.Text = previewText;
            // The profile name is already present in the tooltip and log. Keep
            // the inline readback compact so it remains aligned with the value
            // column without being truncated on the narrower settings layout.
            block.NicItrStatusLabel.Text = "current: test profile";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            if (block.NicItrTimeLabel is not null)
            {
                block.NicItrTimeLabel.Text = FormatNicItrQueueDetailText(previewValues, profile);
                block.NicItrTimeLabel.ForeColor = _mutedText;
            }

            SetNicItrTooltip(block, $"{profile.FamilyName}\nTest preview only. No driver read or write.");
            WriteLog($"NIC.ITR.TEST: {block.Device.InstanceId} profile=\"{profile.FamilyName}\" preview={previewText}");
            return;
        }

        int generation = ++block.NicItrOperationGeneration;
        if (showReadingStatus)
        {
            block.NicItrStatusLabel.Text = "current: reading...";
            if (block.NicItrTimeLabel is not null)
            {
                block.NicItrTimeLabel.Text = "time: reading...";
                block.NicItrTimeLabel.ForeColor = _statusInactive;
            }
            block.NicItrStatusLabel.ForeColor = _statusInactive;
            SetNicItrTooltip(block, $"{profile.FamilyName}\nreading...");
        }

        string instanceId = block.Device.InstanceId;
        try
        {
            (bool ok, List<ulong> values, string? error) result = await Task.Run(() =>
            {
                bool ok = TryReadNicItr(instanceId, profile, out List<ulong> values, out string? error);
                return (ok, values, error);
            });

            if (IsDisposed
                || block.NicItrBox.IsDisposed
                || block.NicItrStatusLabel.IsDisposed
                || generation != block.NicItrOperationGeneration)
            {
                return;
            }

            if (!result.ok)
            {
                bool needsCheck = result.error?.Contains("press CHECK", StringComparison.OrdinalIgnoreCase) == true;
                block.NicItrStatusLabel.Text = $"current: {FormatNicItrError(result.error)}";
                block.NicItrStatusLabel.ForeColor = needsCheck
                    ? _statusInactive
                    : (IsNicItrActionableError(result.error) || IsNicItrDriverLoadError(result.error) ? _statusDanger : _mutedText);
                if (block.NicItrTimeLabel is not null)
                {
                    block.NicItrTimeLabel.Text = "time: unavailable";
                    block.NicItrTimeLabel.ForeColor = _statusInactive;
                }
                SetNicItrTooltip(block, $"{profile.FamilyName}\nread failed: {result.error}");
                WriteLog(QaRun.IsReadOnly && result.error?.Contains("QA read-only session:", StringComparison.Ordinal) == true
                    ? $"NIC.ITR.READ.SKIPPED: {instanceId} reason=qa-read-only hardwareValidation=SKIPPED"
                    : IsImodNeedsCheckStatus(result.error)
                    ? $"NIC.ITR.READ.SKIPPED: {instanceId} reason=driver-not-loaded action=press-CHECK"
                    : $"NIC.ITR.READ.FAILED: {instanceId} error=\"{SanitizeLogValue(result.error)}\"");
                return;
            }

            string valueText = FormatNicItrValueList(result.values, profile);
            string summaryText = "current: read successfully";
            string detailText = FormatNicItrQueueDetailText(result.values, profile);
            string timingText = FormatNicItrTimingList(result.values, profile);
            block.NicItrBox.Text = valueText;
            block.NicItrStatusLabel.Text = summaryText;
            block.NicItrStatusLabel.ForeColor = _statusActive;
            if (block.NicItrTimeLabel is not null)
            {
                block.NicItrTimeLabel.Text = detailText;
                block.NicItrTimeLabel.ForeColor = _mutedText;
            }
            SetNicItrTooltip(block, $"{profile.FamilyName}\nraw: {valueText}\ntime: {timingText}");
            WriteLog($"NIC.ITR.READ: {instanceId} profile=\"{profile.FamilyName}\" values={valueText} timing=\"{timingText}\"");
        }
        catch (Exception ex)
        {
            WriteLog($"NIC.ITR.READ: {instanceId} exception: {ex.Message}");
            if (IsDisposed
                || block.NicItrStatusLabel is null
                || block.NicItrStatusLabel.IsDisposed
                || generation != block.NicItrOperationGeneration)
            {
                return;
            }

            block.NicItrStatusLabel.Text = "current: read failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            if (block.NicItrTimeLabel is not null && !block.NicItrTimeLabel.IsDisposed)
            {
                block.NicItrTimeLabel.Text = "time: unavailable";
                block.NicItrTimeLabel.ForeColor = _statusInactive;
            }

            SetNicItrTooltip(block, $"{profile.FamilyName}\nread failed: {ex.Message}");
        }
    }

    private void CheckImodDriverFromNicBlock(DeviceBlock block)
    {
        WriteLog($"UI: CHECK NIC ITR button clicked device={block.Device.InstanceId}");
        if (block.Device.IsTestDevice)
        {
            if (block.NicItrStatusLabel is not null)
            {
                block.NicItrStatusLabel.Text = "current: driver check skipped (test)";
                block.NicItrStatusLabel.ForeColor = _statusInactive;
            }

            return;
        }

        if (TryBlockBusyHardwareWrite("NIC ITR CHECK") || TryBlockSandboxHardwareWrite("NIC ITR CHECK"))
        {
            return;
        }

        if (!TryCheckImodDriver(out string? error))
        {
            if (block.NicItrStatusLabel is not null)
            {
                block.NicItrStatusLabel.Text = "current: driver load failed";
                block.NicItrStatusLabel.ForeColor = _statusDanger;
                SetNicItrTooltip(block, "DTIMOD driver unavailable. Open the session log for diagnostics.");
            }

            WriteLog($"NIC.ITR.CHECK: failed device={block.Device.InstanceId} error={error}");
            OperationReport report = new();
            report.MarkNoChangesMade();
            report.AddError(
                "NIC ITR DRIVER",
                FormatImodUnavailableUserMessage(error, includeNotChanged: false),
                error ?? "Unknown DTIMOD driver load error.");
            ShowOperationResult(
                report,
                string.Empty,
                "The driver check did not complete.",
                operationName: "NIC ITR CHECK");
            MaybeOfferVulnerableDriverBlocklistDisable(report);
            return;
        }

        WriteLog($"NIC.ITR.CHECK: ok device={block.Device.InstanceId}");
        RefreshNicItrBlock(block, showReadingStatus: false);
    }

    private async void ApplyNicItrFromBlock(DeviceBlock block)
    {
        if (block.NicItrBox is null || block.NicItrStatusLabel is null)
        {
            return;
        }

        NicItrProfile? profile = TryGetNicItrProfile(block.Device.InstanceId);
        if (profile is null)
        {
            block.NicItrStatusLabel.Text = "current: unsupported";
            block.NicItrStatusLabel.ForeColor = _statusInactive;
            SetNicItrTooltip(block, "NIC ITR is unsupported for this adapter.");
            return;
        }

        if (!TryParseNicItrInput(block.NicItrBox.Text ?? string.Empty, profile, out List<ulong> values))
        {
            block.NicItrStatusLabel.Text = "current: invalid input";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            WriteLog($"NIC.ITR.WRITE: {block.Device.InstanceId} invalid input=\"{block.NicItrBox.Text}\"");
            return;
        }

        NormalizeNicItrTextBox(block, profile, values);
        if (block.Device.IsTestDevice)
        {
            EnsureTestDeviceState(block.Device).NicItrValue = FormatNicItrValueList(values, profile);
            block.NicItrStatusLabel.Text = "current: test preview";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.TEST.WRITE.SKIP: {block.Device.InstanceId} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)}");
            return;
        }

        if (TryBlockBusyHardwareWrite("NIC ITR SET") || TryBlockSandboxHardwareWrite("NIC ITR SET"))
        {
            return;
        }

        if (!CreateDeviceTweakerBackup("pre-nic-itr", showDialog: false))
        {
            block.NicItrStatusLabel.Text = "current: backup failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            SetNicItrTooltip(block, $"{profile.FamilyName}\nwrite cancelled: automatic backup failed");
            WriteLog($"NIC.ITR.WRITE: {block.Device.InstanceId} cancelled because automatic backup failed");
            OperationReport report = new();
            report.MarkNoChangesMade();
            report.AddError("AUTOMATIC BACKUP", "The backup could not be created. No changes were made.", "Pre-NIC-ITR backup failed.");
            ShowOperationResult(
                report,
                string.Empty,
                "NIC ITR was cancelled before any changes were made.",
                operationName: "NIC ITR");
            return;
        }

        int generation = ++block.NicItrOperationGeneration;
        block.NicItrStatusLabel.Text = "current: applying...";
        block.NicItrStatusLabel.ForeColor = _statusInactive;
        if (block.NicItrApplyButton is not null)
        {
            block.NicItrApplyButton.Enabled = false;
        }

        string instanceId = block.Device.InstanceId;
        try
        {
            (bool ok, string? error) result = await Task.Run(() =>
            {
                bool ok = TryWriteNicItr(instanceId, profile, values, out string? error);
                return (ok, error);
            });

            if (IsDisposed
                || block.NicItrBox.IsDisposed
                || block.NicItrStatusLabel.IsDisposed
                || generation != block.NicItrOperationGeneration)
            {
                return;
            }

            if (!result.ok)
            {
                block.NicItrStatusLabel.Text = $"current: {FormatNicItrError(result.error)}";
                block.NicItrStatusLabel.ForeColor = IsNicItrActionableError(result.error) || IsNicItrDriverLoadError(result.error) ? _statusDanger : _mutedText;
                SetNicItrTooltip(block, $"{profile.FamilyName}\nwrite failed: {result.error}");
                WriteLog($"NIC.ITR.WRITE: {instanceId} failed: {result.error}");
                return;
            }

            RefreshNicItrBlock(block);
        }
        catch (Exception ex)
        {
            WriteLog($"NIC.ITR.WRITE: {instanceId} exception: {ex.Message}");
            if (IsDisposed
                || block.NicItrStatusLabel is null
                || block.NicItrStatusLabel.IsDisposed
                || generation != block.NicItrOperationGeneration)
            {
                return;
            }

            block.NicItrStatusLabel.Text = "current: write failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            SetNicItrTooltip(block, $"{profile.FamilyName}\nwrite failed: {ex.Message}");
        }
        finally
        {
            if (!IsDisposed && block.NicItrApplyButton is not null && !block.NicItrApplyButton.IsDisposed)
            {
                block.NicItrApplyButton.Enabled = true;
            }
        }
    }

    private void SaveNicItrPersistenceFromBlock(DeviceBlock block)
    {
        if (block.NicItrBox is null || block.NicItrStatusLabel is null)
        {
            return;
        }

        NicItrProfile? profile = TryGetNicItrProfile(block.Device.InstanceId);
        if (profile is null)
        {
            block.NicItrStatusLabel.Text = "current: unsupported";
            block.NicItrStatusLabel.ForeColor = _statusInactive;
            return;
        }

        if (!TryParseNicItrInput(block.NicItrBox.Text ?? string.Empty, profile, out List<ulong> values))
        {
            block.NicItrStatusLabel.Text = $"current: enter 1 or {profile.MaxQueues} values";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.SAVE: {block.Device.InstanceId} invalid input=\"{block.NicItrBox.Text}\"");
            return;
        }

        NormalizeNicItrTextBox(block, profile, values);
        if (block.Device.IsTestDevice)
        {
            EnsureTestDeviceState(block.Device).NicItrValue = FormatNicItrValueList(values, profile);
            block.NicItrStatusLabel.Text = "current: test preview";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.TEST.SAVE.SKIP: {block.Device.InstanceId} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)}");
            return;
        }

        if (TryBlockSandboxHardwareWrite("NIC ITR SAVE") || TryBlockBusyHardwareWrite("NIC ITR SAVE"))
        {
            return;
        }

        if (!CreateDeviceTweakerBackup("pre-nic-itr-save", showDialog: false))
        {
            block.NicItrStatusLabel.Text = "current: backup failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            SetNicItrTooltip(block, $"{profile.FamilyName}\nsave cancelled: automatic backup failed");
            WriteLog($"NIC.ITR.SAVE: {block.Device.InstanceId} cancelled because automatic backup failed");
            return;
        }

        try
        {
            ImodConfig config = ReadSavedStartupConfig(out _);
            string scriptPath = GetStartupProfilePath();
            string hwid = GetNicItrPersistenceKey(block.Device.InstanceId);
            config.NicItrEntries.RemoveAll(e => string.Equals(e.Hwid, hwid, StringComparison.OrdinalIgnoreCase));
            config.NicItrEntries.Add(new NicItrConfigEntry
            {
                Hwid = hwid,
                FamilyName = profile.FamilyName,
                BaseOffset = profile.BaseOffset,
                Stride = profile.Stride,
                Queues = profile.MaxQueues,
                Width = profile.ReadWidth,
                Mask = profile.ReadMask,
                OrBits = profile.WriteOrBits,
                Values = values,
            });

            SaveStartupProfile(config);

            block.NicItrStatusLabel.Text = "current: saved for startup";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.SAVE: {block.Device.InstanceId} key={hwid} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)} path={scriptPath}");

            // Adaptive driver moderation would overwrite the saved register value.
            List<NicDriverModeration.Change> moderation = NicDriverModeration.PlanDisable(ReadNicDriverModeration(block.Device.InstanceId));
            if (!WriteNicDriverModeration(block.Device.InstanceId, moderation, "OFF"))
            {
                block.NicItrStatusLabel.Text = "current: saved; driver moderation unchanged";
                block.NicItrStatusLabel.ForeColor = _statusDanger;
            }

            UpdateBlockInfoText(block);
        }
        catch (Exception ex)
        {
            block.NicItrStatusLabel.Text = "current: save failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            WriteLog($"NIC.ITR.SAVE: {block.Device.InstanceId} failed: {ex.Message}");
        }
    }

    private bool TryReadNicItr(string instanceId, NicItrProfile profile, out List<ulong> values, out string? error)
    {
        values = [];
        error = null;

        if (!IsAdministrator())
        {
            error = "administrator privileges required";
            return false;
        }

        if (!TryGetPciMemoryRangeByInstanceId(instanceId, out ulong baseAddress, out ulong memoryLength, out error)
            || !TryValidateNicItrMemoryRange(profile, memoryLength, out error))
        {
            return false;
        }

        // Opening an existing driver cannot start a loader or upgrade saved settings.
        {
            if (!ImodDriverContext.TryOpenExisting(WriteLog, out ImodDriverContext? driverContext, out error))
            {
                return false;
            }

            using ImodDriverContext ctx = driverContext!;
            for (int q = 0; q < profile.MaxQueues; q++)
            {
                ulong address = baseAddress + profile.BaseOffset + (profile.Stride * (uint)q);
                if (!TryReadNicRegister(ctx, address, profile.ReadWidth, out ulong raw, out error))
                {
                    return false;
                }

                values.Add(raw & profile.ReadMask);
            }

            return true;
        }
    }

    private bool TryWriteNicItr(string instanceId, NicItrProfile profile, IReadOnlyList<ulong> values, out string? error)
    {
        error = null;
        if (values.Count == 0)
        {
            error = "no values";
            return false;
        }

        if (!IsAdministrator())
        {
            error = "administrator privileges required";
            return false;
        }

        if (!TryGetPciMemoryRangeByInstanceId(instanceId, out ulong baseAddress, out ulong memoryLength, out error)
            || !TryValidateNicItrMemoryRange(profile, memoryLength, out error))
        {
            return false;
        }

        bool persistDriver = ShouldPersistSharedImodDriver();
        if (!EnsureImodDriverOnDisk(persistDriver, out string driverPath, out error))
        {
            return false;
        }

        try
        {
            if (!ImodDriverContext.TryInitialize(driverPath, WriteLog, out ImodDriverContext? driverContext, out error))
            {
                LogImodDriverLoadDiagnostics(driverPath, error);
                return false;
            }

            using ImodDriverContext ctx = driverContext!;
            if (!TryProgramNicItrRegisters(profile, baseAddress, values, new DriverNicItrRegisters(ctx), out error))
            {
                return false;
            }

            WriteLog($"NIC.ITR.WRITE: {instanceId} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)}");
            return true;
        }
        finally
        {
            if (!persistDriver && !IsImodDriverSystemPath(driverPath))
            {
                DeleteFileIfExists(driverPath, "IMOD.DRIVER");
            }
        }
    }

    private sealed class DriverNicItrRegisters(ImodDriverContext ctx) : INicItrRegisters
    {
        public bool TryWrite(ulong address, int width, ulong value, out string? error)
            => TryWriteNicRegister(ctx, address, width, value, out error);
        public bool TryRead(ulong address, int width, out ulong value, out string? error)
            => TryReadNicRegister(ctx, address, width, out value, out error);
    }

    // The production write/readback sequence is shared with the in-memory QA bus.
    internal static bool TryProgramNicItrRegisters(
        NicItrProfile profile, ulong baseAddress, IReadOnlyList<ulong> values,
        INicItrRegisters registers, out string? error)
        => NicItrService.TryProgramNicItrRegisters(profile, baseAddress, values, registers, out error);

    private bool ShouldPersistSharedImodDriver()
    {
        EnsureImodConfigLoaded();
        return _imodConfigCache is not null
            && (_imodConfigCache.HasScript || HasCustomImod(_imodConfigCache));
    }

    private static bool TryReadNicRegister(ImodDriverContext ctx, ulong address, int width, out ulong value, out string? error)
    {
        value = 0;
        if (width == 16)
        {
            if (!TryReadPhys16(ctx, address, out ushort word, out error))
            {
                return false;
            }

            value = word;
            return true;
        }

        if (!TryReadPhys32(ctx, address, out uint dword, out error))
        {
            return false;
        }

        value = dword;
        return true;
    }

    private static bool TryWriteNicRegister(ImodDriverContext ctx, ulong address, int width, ulong value, out string? error)
    {
        return width == 16
            ? TryWritePhys16(ctx, address, (ushort)value, out error)
            : TryWritePhys32(ctx, address, (uint)value, out error);
    }

    private static bool TryReadPhys16(ImodDriverContext ctx, ulong address, out ushort value, out string? error)
    {
        value = 0;

        if (!TryReadPhysicalMemory(ctx, address, 2, out ulong raw, out error))
        {
            return false;
        }

        value = unchecked((ushort)raw);
        return true;
    }

    private static bool TryWritePhys16(ImodDriverContext ctx, ulong address, ushort value, out string? error)
    {
        return TryWritePhysicalMemory(ctx, address, 2, value, out error);
    }

    private static bool TryValidateNicItrMemoryRange(NicItrProfile profile, ulong memoryLength, out string? error)
        => NicItrService.TryValidateNicItrMemoryRange(profile, memoryLength, out error);

    private bool TryGetPciMemoryRangeByInstanceId(string instanceId, out ulong baseAddress, out ulong memoryLength, out string? error)
    {
        baseAddress = 0;
        memoryLength = 0;
        error = null;
        string target = NormalizeInstanceId(instanceId);

        IntPtr devInfoSet = SetupDiGetClassDevsW(IntPtr.Zero, "PCI", IntPtr.Zero, DigcfPresent | DigcfAllClasses);
        if (devInfoSet == InvalidHandleValue)
        {
            error = $"failed to enumerate PCI devices: {GetWin32ErrorMessage(Marshal.GetLastWin32Error())}";
            return false;
        }

        try
        {
            for (uint index = 0; ; index++)
            {
                SP_DEVINFO_DATA devInfo = new()
                {
                    cbSize = (uint)Marshal.SizeOf<SP_DEVINFO_DATA>(),
                };

                if (!SetupDiEnumDeviceInfo(devInfoSet, index, ref devInfo))
                {
                    int lastError = Marshal.GetLastWin32Error();
                    if (lastError == ErrorNoMoreItems)
                    {
                        break;
                    }

                    error = $"failed to enumerate device info: {GetWin32ErrorMessage(lastError)}";
                    return false;
                }

                if (!TryGetDeviceInstanceId(devInfoSet, ref devInfo, out string currentId)
                    || !string.Equals(NormalizeInstanceId(currentId), target, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!TryGetDeviceMemoryRange(devInfo.DevInst, out baseAddress, out memoryLength, out error))
                {
                    return false;
                }

                return true;
            }
        }
        finally
        {
            _ = SetupDiDestroyDeviceInfoList(devInfoSet);
        }

        error = "PCI device not found";
        return false;
    }

    private static bool TryGetPciVenDev(string instanceId, out string vendorId, out string deviceId)
        => NicItrService.TryGetPciVenDev(instanceId, out vendorId, out deviceId);

    private static bool TryParseNicItrInput(string text, NicItrProfile profile, out List<ulong> values)
        => NicItrService.TryParseNicItrInput(text, profile, out values);

    private void UpdateNicItrInputTimeLabel(DeviceBlock block)
    {
        if (block.NicItrTimeLabel is null)
        {
            return;
        }

        NicItrProfile? profile = TryGetNicItrProfile(block.Device.InstanceId);
        if (profile is null)
        {
            block.NicItrTimeLabel.Text = "time: unsupported";
            block.NicItrTimeLabel.ForeColor = _statusInactive;
            return;
        }

        if (!TryParseNicItrInput(block.NicItrBox?.Text ?? string.Empty, profile, out List<ulong> values))
        {
            block.NicItrTimeLabel.Text = $"time: enter 1 or {profile.MaxQueues} values";
            block.NicItrTimeLabel.ForeColor = _statusDanger;
            return;
        }

        block.NicItrTimeLabel.Text = values.Count == 1 && profile.MaxQueues > 1
            ? $"preview: all queues -> {FormatNicItrTimingDetail(values[0], profile)}"
            : FormatNicItrQueueDetailText(values, profile, preview: true);
        block.NicItrTimeLabel.ForeColor = _mutedText;
    }

    private static void NormalizeNicItrTextBox(DeviceBlock block, NicItrProfile profile, IReadOnlyList<ulong> values)
    {
        if (block.NicItrBox is null)
        {
            return;
        }

        string normalized = FormatNicItrValueList(values, profile);
        if (!string.Equals(block.NicItrBox.Text?.Trim(), normalized, StringComparison.OrdinalIgnoreCase))
        {
            block.NicItrBox.Text = normalized;
        }
    }

    private string GetNicItrPersistenceKey(string instanceId)
    {
        Match match = Regex.Match(
            instanceId ?? string.Empty,
            "VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})",
            RegexOptions.CultureInvariant);
        return match.Success
            ? $"VEN_{match.Groups[1].Value.ToUpperInvariant()}&DEV_{match.Groups[2].Value.ToUpperInvariant()}"
            : NormalizeInstanceId(instanceId);
    }

    private static bool TryParseUInt64Flexible(string text, out ulong value)
        => NicItrService.TryParseUInt64Flexible(text, out value);

    private static string FormatNicItrValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
        => NicItrService.FormatNicItrValueList(values, profile);

    private static string FormatNicItrValue(ulong value, NicItrProfile profile)
        => NicItrService.FormatNicItrValue(value, profile);

    private static string FormatNicItrLabeledValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
        => NicItrService.FormatNicItrLabeledValueList(values, profile);

    private static string GetNicItrVectorLabel(NicItrProfile profile, int index)
        => NicItrService.GetNicItrVectorLabel(profile, index);

    private static string FormatNicItrQueueSummary(IReadOnlyList<ulong> values, NicItrProfile profile)
        => NicItrService.FormatNicItrQueueSummary(values, profile);

    private static string FormatNicItrQueueDetailText(IReadOnlyList<ulong> values, NicItrProfile profile, bool preview = false)
        => NicItrService.FormatNicItrQueueDetailText(values, profile, preview);

    private static string FormatNicItrTimingList(IReadOnlyList<ulong> values, NicItrProfile profile)
        => NicItrService.FormatNicItrTimingList(values, profile);

    private static string FormatNicItrTiming(ulong raw, NicItrProfile profile)
        => NicItrService.FormatNicItrTiming(raw, profile);

    private static string FormatNicItrTimingDetail(ulong raw, NicItrProfile profile)
        => NicItrService.FormatNicItrTimingDetail(raw, profile);

    private static string FormatDurationUs(ulong us)
        => NicItrService.FormatDurationUs(us);

    private static string FormatDurationNs(ulong ns)
        => NicItrService.FormatDurationNs(ns);

    private static string FormatRealtekIntrMit(ulong raw, uint timerUnitUs, bool extended)
        => NicItrService.FormatRealtekIntrMit(raw, timerUnitUs, extended);

    private static string FormatRealtekIntrMitDetail(ulong raw, uint timerUnitUs, bool extended)
        => NicItrService.FormatRealtekIntrMitDetail(raw, timerUnitUs, extended);

    private void SetNicItrTooltip(DeviceBlock block, string text)
    {
        try
        {
            if (block.NicItrStatusLabel is not null)
            {
                _copyToolTip.SetToolTip(block.NicItrStatusLabel, text);
            }

            if (block.NicItrBox is not null)
            {
                _copyToolTip.SetToolTip(block.NicItrBox, text);
            }

            if (block.NicItrTimeLabel is not null)
            {
                _copyToolTip.SetToolTip(block.NicItrTimeLabel, text);
            }
        }
        catch
        {
        }
    }

    private static bool IsNicItrActionableError(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return false;
        }

        return error.Contains("administrator", StringComparison.OrdinalIgnoreCase)
            || error.Contains("админист", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNicItrDriverLoadError(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return false;
        }

        return error.Contains("IMOD driver service", StringComparison.OrdinalIgnoreCase)
            || error.Contains("driver service", StringComparison.OrdinalIgnoreCase)
            || error.Contains("DeviceTweakerImod", StringComparison.OrdinalIgnoreCase)
            || error.Contains("start IMOD driver", StringComparison.OrdinalIgnoreCase)
            || error.Contains("driver not loaded", StringComparison.OrdinalIgnoreCase)
            || error.Contains("press CHECK", StringComparison.OrdinalIgnoreCase)
            || error.Contains("digital signature", StringComparison.OrdinalIgnoreCase)
            || error.Contains("подпис", StringComparison.OrdinalIgnoreCase)
            || error.Contains("цифров", StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatNicItrError(string? error)
    {
        if (string.IsNullOrWhiteSpace(error))
        {
            return "unavailable";
        }

        if (IsNicItrDriverLoadError(error))
        {
            if (error.Contains("press CHECK", StringComparison.OrdinalIgnoreCase))
            {
                return "press CHECK";
            }

            return IsImodSignatureRejectedError(error) ? "kernel CI blocked" : "driver not loaded";
        }

        if (error.Contains("administrator", StringComparison.OrdinalIgnoreCase)
            || error.Contains("админист", StringComparison.OrdinalIgnoreCase))
        {
            return "admin required";
        }

        return "unavailable";
    }
}
