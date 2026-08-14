using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private enum NicItrTimingKind
    {
        None,
        IntelEitr,
        IntelItr,
        RealtekIntrMit,
        RealtekIntrMitV2,
    }

    private sealed record NicItrProfile(
        string FamilyName,
        string VendorId,
        string[] DeviceIds,
        uint BaseOffset,
        uint Stride,
        int MaxQueues,
        int ReadWidth,
        ulong ReadMask,
        ulong WriteOrBits,
        NicItrTimingKind TimingKind,
        string[]? VectorLabels = null);

    private static readonly NicItrProfile[] NicItrProfiles =
    [
        new("Intel I225/I226 (EITR)", "8086", ["15F2", "15F3", "0D9F", "5502", "125B", "125C", "125D", "5503"], 0x1680, 0x4, 5, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, ["Other", "Q0", "Q1", "Q2", "Q3"]),
        new("Intel I210/I211 (EITR)", "8086", ["1533", "1536", "1537", "1538", "1539", "157B", "157C", "1F40", "1F41", "1F45"], 0x1680, 0x4, 4, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Intel I350 (EITR)", "8086", ["1521", "1522", "1523", "1524"], 0x1680, 0x4, 8, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Intel 82580 (EITR)", "8086", ["150E", "150F", "1510", "1511"], 0x1680, 0x4, 10, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Intel 82576 (EITR)", "8086", ["1516", "1518", "1526"], 0x1680, 0x4, 25, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr),
        new("Killer E3100 (EITR)", "8086", ["3100", "3101", "3102"], 0x1680, 0x4, 5, 32, 0x00007FFC, 0x80000000, NicItrTimingKind.IntelEitr, ["Other", "Q0", "Q1", "Q2", "Q3"]),
        new("Intel I219 (ITR)", "8086", ["15B7", "15B8", "15B9", "15D7", "15D8", "15E3", "15BB", "15BC", "15BD", "15BE", "0D4C", "0D4D", "0D4E", "0D4F", "0D53", "0D55", "0D5C", "0D5D", "0D5E", "0D5F", "15FB", "15FC", "1A1E", "1A1F", "550A", "550B", "550C", "550D", "550E", "550F", "0DC5", "0DC6", "0DC7", "0DC8", "1A1C", "1A1D", "15F9", "15FA", "3166", "3167", "3197", "3198", "4DF4", "4B33", "4B34", "4DC2", "4DC3", "54B4", "54B5", "54B6", "54B7", "0126", "153A", "153B", "1559", "155A", "156F", "1570", "15D6"], 0x00C4, 0x0, 1, 32, 0x0000FFFF, 0, NicItrTimingKind.IntelItr),
        new("Realtek RTL8111/8168", "10EC", ["8168", "8161", "8136", "8167", "8169"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
        new("Realtek RTL8125/8126", "10EC", ["8125", "8162", "8126"], 0x0A00, 0x8, 4, 32, 0x7F7F7F7F, 0, NicItrTimingKind.RealtekIntrMitV2),
        new("Realtek RTL8168KB", "10EC", ["3000"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
        new("Killer E2500/E2600", "10EC", ["2600", "2502", "2500"], 0x00E2, 0x0, 1, 16, 0xFFFF, 0, NicItrTimingKind.RealtekIntrMit),
    ];

    private static NicItrProfile? TryGetNicItrProfile(string instanceId)
    {
        if (!TryGetPciVenDev(instanceId, out string ven, out string dev))
        {
            return null;
        }

        foreach (NicItrProfile profile in NicItrProfiles)
        {
            if (string.Equals(profile.VendorId, ven, StringComparison.OrdinalIgnoreCase)
                && profile.DeviceIds.Any(id => string.Equals(id, dev, StringComparison.OrdinalIgnoreCase)))
            {
                return profile;
            }
        }

        return null;
    }

    private async void RefreshNicItrBlock(DeviceBlock block)
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
            List<ulong> previewValues = Enumerable.Repeat(0UL, Math.Max(1, profile.MaxQueues)).ToList();
            string previewText = FormatNicItrValueList(previewValues, profile);
            block.NicItrBox.Text = previewText;
            block.NicItrStatusLabel.Text = $"current: test profile {profile.FamilyName}";
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
        block.NicItrStatusLabel.Text = "current: reading...";
        if (block.NicItrTimeLabel is not null)
        {
            block.NicItrTimeLabel.Text = "time: reading...";
            block.NicItrTimeLabel.ForeColor = _statusInactive;
        }
        block.NicItrStatusLabel.ForeColor = _statusInactive;
        SetNicItrTooltip(block, $"{profile.FamilyName}\nreading...");

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
                WriteLog($"NIC.ITR.READ: {instanceId} failed: {result.error}");
                return;
            }

            string valueText = FormatNicItrValueList(result.values, profile);
            string summaryText = FormatNicItrQueueSummary(result.values, profile);
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

        if (TryBlockSandboxHardwareWrite("NIC ITR CHECK"))
        {
            return;
        }

        if (block.NicItrStatusLabel is not null)
        {
            block.NicItrStatusLabel.Text = "current: loading driver...";
            block.NicItrStatusLabel.ForeColor = _statusInactive;
        }

        if (!TryCheckImodDriver(out string? error))
        {
            if (block.NicItrStatusLabel is not null)
            {
                block.NicItrStatusLabel.Text = "current: driver load failed";
                block.NicItrStatusLabel.ForeColor = _statusDanger;
                SetNicItrTooltip(block, error ?? "DTIMOD.sys load failed");
            }

            WriteLog($"NIC.ITR.CHECK: failed device={block.Device.InstanceId} error={error}");
            ShowThemedInfo($"IMOD driver load failed.\n{error}");
            return;
        }

        WriteLog($"NIC.ITR.CHECK: ok device={block.Device.InstanceId}");
        RefreshNicItrBlock(block);
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
            block.NicItrStatusLabel.Text = "current: test preview";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.TEST.WRITE.SKIP: {block.Device.InstanceId} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)}");
            return;
        }

        if (TryBlockSandboxHardwareWrite("NIC ITR SET"))
        {
            return;
        }

        if (!CreateDeviceTweakerBackup("pre-nic-itr", showDialog: false))
        {
            block.NicItrStatusLabel.Text = "current: backup failed";
            block.NicItrStatusLabel.ForeColor = _statusDanger;
            SetNicItrTooltip(block, $"{profile.FamilyName}\nwrite cancelled: automatic backup failed");
            WriteLog($"NIC.ITR.WRITE: {block.Device.InstanceId} cancelled because automatic backup failed");
            ShowThemedInfo("NIC ITR apply was cancelled because the automatic backup failed.\nNo registry/hardware changes were made.");
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
            block.NicItrStatusLabel.Text = "current: test preview";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.TEST.SAVE.SKIP: {block.Device.InstanceId} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)}");
            return;
        }

        try
        {
            ResolveImodPaths(out string? scriptPath);
            if (string.IsNullOrWhiteSpace(scriptPath))
            {
                scriptPath = GetImodStartupPath();
            }

            bool scriptExists = File.Exists(scriptPath);
            ImodConfig config = scriptExists ? ParseImodScriptFile(scriptPath) : new ImodConfig();
            config.HasScript = scriptExists;

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

            WriteImodScript(config, scriptPath);
            config.HasScript = true;
            _imodConfigCache = config;
            _imodScriptPath = scriptPath;
            _imodConfigLoaded = true;

            block.NicItrStatusLabel.Text = "current: saved for startup";
            block.NicItrStatusLabel.ForeColor = _statusActive;
            UpdateNicItrInputTimeLabel(block);
            WriteLog($"NIC.ITR.SAVE: {block.Device.InstanceId} key={hwid} profile=\"{profile.FamilyName}\" values={FormatNicItrValueList(values, profile)} path={scriptPath}");
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

        if (!TryGetPciMemoryBaseByInstanceId(instanceId, out ulong baseAddress, out error))
        {
            return false;
        }

        bool persistDriver = ShouldPersistSharedImodDriver();
        if (!EnsureImodDriverOnDisk(persistDriver, out string driverPath, out error))
        {
            return false;
        }

        if (!IsImodDriverAlreadyAvailable())
        {
            error = "driver not loaded (press CHECK)";
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
        finally
        {
            if (!persistDriver && !IsImodDriverSystemPath(driverPath))
            {
                DeleteFileIfExists(driverPath, "IMOD.DRIVER");
            }
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

        if (!TryGetPciMemoryBaseByInstanceId(instanceId, out ulong baseAddress, out error))
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
            for (int q = 0; q < profile.MaxQueues; q++)
            {
                ulong selected = q < values.Count ? values[q] : values[0];
                ulong finalValue = (selected & profile.ReadMask) | profile.WriteOrBits;
                ulong address = baseAddress + profile.BaseOffset + (profile.Stride * (uint)q);
                if (!TryWriteNicRegister(ctx, address, profile.ReadWidth, finalValue, out error))
                {
                    return false;
                }
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

    private bool TryGetPciMemoryBaseByInstanceId(string instanceId, out ulong baseAddress, out string? error)
    {
        baseAddress = 0;
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

                if (!TryGetDeviceMemoryBase(devInfo.DevInst, out baseAddress, out error))
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
    {
        vendorId = string.Empty;
        deviceId = string.Empty;
        Match match = Regex.Match(
            instanceId ?? string.Empty,
            "VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})",
            RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return false;
        }

        vendorId = match.Groups[1].Value.ToUpperInvariant();
        deviceId = match.Groups[2].Value.ToUpperInvariant();
        return true;
    }

    private static bool TryParseNicItrInput(string text, NicItrProfile profile, out List<ulong> values)
    {
        values = [];
        string[] parts = (text ?? string.Empty)
            .Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || (parts.Length != 1 && parts.Length != profile.MaxQueues))
        {
            return false;
        }

        foreach (string part in parts)
        {
            if (!TryParseUInt64Flexible(part, out ulong parsed))
            {
                values.Clear();
                return false;
            }

            values.Add(parsed & profile.ReadMask);
        }

        return values.Count > 0;
    }

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
    {
        value = 0;
        string trimmed = text.Trim();
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            return ulong.TryParse(trimmed[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value);
        }

        return ulong.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private static string FormatNicItrValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        string[] formatted = values.Select(v => FormatNicItrValue(v, profile)).ToArray();
        return string.Join(", ", formatted);
    }

    private static string FormatNicItrValue(ulong value, NicItrProfile profile)
    {
        value &= profile.ReadMask;
        if (value == 0)
        {
            return "0x0";
        }

        int digits = profile.ReadWidth == 16 ? 4 : 8;
        return "0x" + value.ToString($"X{digits}", CultureInfo.InvariantCulture);
    }

    private static string FormatNicItrLabeledValueList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "-";
        }

        string[] formatted = values
            .Select((value, index) => $"{GetNicItrVectorLabel(profile, index)}={FormatNicItrValue(value, profile)}")
            .ToArray();
        return string.Join(", ", formatted);
    }

    private static string GetNicItrVectorLabel(NicItrProfile profile, int index)
    {
        if (profile.VectorLabels is not null && index >= 0 && index < profile.VectorLabels.Length)
        {
            return profile.VectorLabels[index];
        }

        return profile.MaxQueues == 1 ? "ITR" : $"Q{index}";
    }

    private static string FormatNicItrQueueSummary(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0)
        {
            return "current: -";
        }

        if (profile.MaxQueues == 1 || values.Count == 1)
        {
            return $"current: {FormatNicItrValue(values[0], profile)} | {FormatNicItrTimingDetail(values[0], profile)}";
        }

        string[] states = values
            .Select((value, index) =>
            {
                string state = (value & profile.ReadMask) == 0 ? "off" : "active";
                return $"{GetNicItrVectorLabel(profile, index)} {state}";
            })
            .ToArray();

        return "current: " + string.Join(" | ", states);
    }

    private static string FormatNicItrQueueDetailText(IReadOnlyList<ulong> values, NicItrProfile profile, bool preview = false)
    {
        if (values.Count == 0)
        {
            return preview ? "preview: -" : "queues: -";
        }

        if (values.Count > 4)
        {
            string prefix = preview ? "preview: " : "queues: ";
            return prefix + FormatNicItrTimingList(values, profile);
        }

        string[] rows = values
            .Select((value, index) =>
            {
                string label = GetNicItrVectorLabel(profile, index);
                string raw = FormatNicItrValue(value, profile).PadRight(profile.ReadWidth == 16 ? 6 : 10);
                return $"{label}: {raw}   {FormatNicItrTimingDetail(value, profile)}";
            })
            .ToArray();

        return string.Join(Environment.NewLine, rows);
    }

    private static string FormatNicItrTimingList(IReadOnlyList<ulong> values, NicItrProfile profile)
    {
        if (values.Count == 0 || profile.TimingKind == NicItrTimingKind.None)
        {
            return "n/a";
        }

        string[] formatted = values
            .Select((value, index) => $"{GetNicItrVectorLabel(profile, index)}={FormatNicItrTiming(value, profile)}")
            .ToArray();
        return string.Join(", ", formatted);
    }

    private static string FormatNicItrTiming(ulong raw, NicItrProfile profile)
    {
        raw &= profile.ReadMask;
        return profile.TimingKind switch
        {
            NicItrTimingKind.IntelEitr => FormatDurationUs(((raw >> 2) & 0x1FFF) * 2),
            NicItrTimingKind.IntelItr => FormatDurationNs(raw * 256),
            NicItrTimingKind.RealtekIntrMit => FormatRealtekIntrMit(raw, timerUnitUs: 125, extended: false),
            NicItrTimingKind.RealtekIntrMitV2 => FormatRealtekIntrMit(raw, timerUnitUs: 1, extended: true),
            _ => "n/a",
        };
    }

    private static string FormatNicItrTimingDetail(ulong raw, NicItrProfile profile)
    {
        raw &= profile.ReadMask;
        if (raw == 0)
        {
            return "Off";
        }

        return profile.TimingKind switch
        {
            NicItrTimingKind.RealtekIntrMit => FormatRealtekIntrMitDetail(raw, timerUnitUs: 125, extended: false),
            NicItrTimingKind.RealtekIntrMitV2 => FormatRealtekIntrMitDetail(raw, timerUnitUs: 1, extended: true),
            NicItrTimingKind.IntelEitr => "Delay " + FormatNicItrTiming(raw, profile),
            NicItrTimingKind.IntelItr => "Delay " + FormatNicItrTiming(raw, profile),
            _ => FormatNicItrTiming(raw, profile),
        };
    }

    private static string FormatDurationUs(ulong us)
    {
        if (us == 0)
        {
            return "Off";
        }

        return us >= 1000
            ? (us / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + " ms"
            : us.ToString(CultureInfo.InvariantCulture) + " us";
    }

    private static string FormatDurationNs(ulong ns)
    {
        if (ns == 0)
        {
            return "Off";
        }

        if (ns >= 1_000_000)
        {
            return (ns / 1_000_000.0).ToString("0.###", CultureInfo.InvariantCulture) + " ms";
        }

        if (ns >= 1000)
        {
            return (ns / 1000.0).ToString("0.###", CultureInfo.InvariantCulture) + " us";
        }

        return ns.ToString(CultureInfo.InvariantCulture) + " ns";
    }

    private static string FormatRealtekIntrMit(ulong raw, uint timerUnitUs, bool extended)
    {
        if (raw == 0)
        {
            return "Off";
        }

        ulong timerMask = extended ? 0x7FUL : 0xFUL;
        ulong frameMask = extended ? 0x7FUL : 0xFUL;
        int rxFrameShift = extended ? 8 : 4;
        int txTimerShift = extended ? 16 : 8;
        int txFrameShift = extended ? 24 : 12;

        ulong rxTimer = raw & timerMask;
        ulong rxFrames = (raw >> rxFrameShift) & frameMask;
        ulong txTimer = (raw >> txTimerShift) & timerMask;
        ulong txFrames = (raw >> txFrameShift) & frameMask;

        string rx = $"{rxTimer * timerUnitUs}us/{rxFrames}f";
        string tx = $"{txTimer * timerUnitUs}us/{txFrames}f";
        return $"Rx:{rx} Tx:{tx}";
    }

    private static string FormatRealtekIntrMitDetail(ulong raw, uint timerUnitUs, bool extended)
    {
        if (raw == 0)
        {
            return "Off";
        }

        ulong timerMask = extended ? 0x7FUL : 0xFUL;
        ulong frameMask = extended ? 0x7FUL : 0xFUL;
        int rxFrameShift = extended ? 8 : 4;
        int txTimerShift = extended ? 16 : 8;
        int txFrameShift = extended ? 24 : 12;

        ulong rxTimer = raw & timerMask;
        ulong rxFrames = (raw >> rxFrameShift) & frameMask;
        ulong txTimer = (raw >> txTimerShift) & timerMask;
        ulong txFrames = (raw >> txFrameShift) & frameMask;

        string rx = $"RX {rxTimer * timerUnitUs}us/{rxFrames}f";
        string tx = $"TX {txTimer * timerUnitUs}us/{txFrames}f";
        return $"{rx}, {tx}";
    }

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
