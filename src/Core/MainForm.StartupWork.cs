namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    // A worker owns this copy: subsequent editor/cache updates cannot change
    // the profile that is being validated and installed.
    internal static ImodConfig SnapshotImodConfig(ImodConfig source)
    {
        ImodConfig copy = new()
        {
            GlobalInterval = source.GlobalInterval,
            GlobalHcsparamsOffset = source.GlobalHcsparamsOffset,
            GlobalRtsoff = source.GlobalRtsoff,
            HasScript = source.HasScript,
            StartupReadError = source.StartupReadError,
        };
        foreach (ImodConfigEntry entry in source.Overrides)
            copy.Overrides.Add(new ImodConfigEntry
            {
                Hwid = entry.Hwid, Enabled = entry.Enabled, Interval = entry.Interval,
                AdaptiveRoleBinding = entry.AdaptiveRoleBinding,
                HcsparamsOffset = entry.HcsparamsOffset, Rtsoff = entry.Rtsoff,
                Intervals = entry.Intervals?.ToList(),
                RoleIntervals = entry.RoleIntervals is null ? null : new(entry.RoleIntervals, entry.RoleIntervals.Comparer),
            });
        foreach (NicItrConfigEntry entry in source.NicItrEntries)
            copy.NicItrEntries.Add(new NicItrConfigEntry
            {
                Hwid = entry.Hwid, FamilyName = entry.FamilyName,
                BaseOffset = entry.BaseOffset, Stride = entry.Stride, Queues = entry.Queues,
                Width = entry.Width, Mask = entry.Mask, OrBits = entry.OrBits,
                Values = entry.Values.ToList(),
            });
        return copy;
    }

    private Task SaveStartupProfile(ImodConfig config)
    {
        // Deliberately synchronous, before taking inputs or scheduling work.
        DemandRealWrite("SaveStartupProfile");
        if (config.StartupReadError is not null)
            throw new InvalidDataException("The startup profile is unreadable; save was cancelled. " + config.StartupReadError);
        ImodConfig snapshot = SnapshotImodConfig(config);
        return SaveStartupProfileCoreAsync(config, snapshot);
    }

    private async Task SaveStartupProfileCoreAsync(ImodConfig owner, ImodConfig snapshot)
    {
        await Task.Run(() =>
        {
            DemandRealWrite("SaveStartupProfile worker");
            // Recheck immediately before replacement, including changes made by
            // another program since the UI first loaded the saved profile.
            _ = ReadSavedStartupConfig(out _);
            InstallStartupBundle(snapshot);
        });
        owner.HasScript = true;
        snapshot.HasScript = true;
        _imodConfigCache = snapshot;
        _imodConfigLoaded = true;
    }

    private Task RemoveImodPersistenceFiles(OperationReport? report = null)
    {
        DemandRealWrite("RemoveImodPersistenceFiles");
        string[] adapterIds = _blocks
            .Where(block => block.Kind is DeviceKind.NET_NDIS or DeviceKind.NET_CX && !block.Device.IsTestDevice)
            .Select(block => block.Device.InstanceId).ToArray();
        return Task.Run(() =>
        {
            DemandRealWrite("RemoveImodPersistenceFiles worker");
            RestoreNicDriverModerationForSavedProfiles(adapterIds, report);
            RemoveStartupBundle(report);
            DeleteFileIfExists(Path.Combine(GetScriptRoot(), "dtimod.sys"), "IMOD.DRIVER.LEGACY", report);
            DeleteFileIfExists(Path.Combine(GetScriptRoot(), ImodDriverName), "IMOD.DRIVER.LEGACY", report);
            WriteLog($"IMOD.DRIVER: keep staged system driver {GetImodDriverSystemPath()}");
        });
    }
}
