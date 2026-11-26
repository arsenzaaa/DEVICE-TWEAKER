namespace DeviceTweakerCS;

// UI draft data only: no controls, registry access or device operations.
internal sealed record DeviceEditorState
{
    internal ulong AffinityMask { get; init; }
    internal int MsiIndex { get; init; }
    internal string LimitText { get; init; } = string.Empty;
    internal int PriorityIndex { get; init; }
    internal int PolicyIndex { get; init; }
    internal bool? PowerSaving { get; init; }
    internal string ImodText { get; init; } = string.Empty;
    internal bool ImodEnabled { get; init; }
    internal bool? RawMouseEnabled { get; init; }
    internal int RawMouseIndex { get; init; } = -1;
    internal int? NdisModeIndex { get; init; }
    internal decimal? RssQueues { get; init; }
    internal string? NicItrText { get; init; }

    internal DeviceSettingChanges GetDeviceChanges(DeviceEditorState current, DeviceKind kind, bool canEditPolicy)
    {
        DeviceSettingChanges changes = DeviceSettingChanges.None;
        if (MsiIndex != current.MsiIndex) changes |= DeviceSettingChanges.Msi;
        if (!string.Equals(LimitText.Trim(), current.LimitText.Trim(), StringComparison.Ordinal))
            changes |= DeviceSettingChanges.MsiLimit;
        if (PriorityIndex != current.PriorityIndex) changes |= DeviceSettingChanges.Priority;
        if (current.PowerSaving.HasValue && PowerSaving != current.PowerSaving) changes |= DeviceSettingChanges.Power;
        bool affinityChanged = AffinityMask != current.AffinityMask
            || (canEditPolicy && PolicyIndex != current.PolicyIndex);
        if (kind == DeviceKind.NET_NDIS)
            affinityChanged |= NdisModeIndex != current.NdisModeIndex || RssQueues != current.RssQueues;
        if (kind != DeviceKind.STOR && affinityChanged) changes |= DeviceSettingChanges.Affinity;
        return changes;
    }

    internal bool HasImodChange(DeviceEditorState current)
        => !string.Equals(ImodText.Trim(), current.ImodText.Trim(), StringComparison.Ordinal);

    internal static DeviceEditorState MergeDraft(DeviceEditorState latest, DeviceEditorState baseline, DeviceEditorState draft)
    {
        static T Pick<T>(T current, T before, T edited) => EqualityComparer<T>.Default.Equals(before, edited) ? current : edited;
        // The mode determines how the processor selection is interpreted. A
        // pending affinity edit must keep its mode when the saved settings change.
        bool affinityEdited = (baseline.AffinityMask, baseline.PolicyIndex, baseline.NdisModeIndex, baseline.RssQueues)
            != (draft.AffinityMask, draft.PolicyIndex, draft.NdisModeIndex, draft.RssQueues);
        DeviceEditorState affinity = affinityEdited ? draft : latest;
        return latest with
        {
            AffinityMask = affinity.AffinityMask,
            MsiIndex = Pick(latest.MsiIndex, baseline.MsiIndex, draft.MsiIndex),
            LimitText = Pick(latest.LimitText, baseline.LimitText, draft.LimitText),
            PriorityIndex = Pick(latest.PriorityIndex, baseline.PriorityIndex, draft.PriorityIndex),
            PolicyIndex = affinity.PolicyIndex,
            PowerSaving = Pick(latest.PowerSaving, baseline.PowerSaving, draft.PowerSaving),
            ImodText = Pick(latest.ImodText, baseline.ImodText, draft.ImodText),
            ImodEnabled = Pick(latest.ImodEnabled, baseline.ImodEnabled, draft.ImodEnabled),
            RawMouseEnabled = Pick(latest.RawMouseEnabled, baseline.RawMouseEnabled, draft.RawMouseEnabled),
            RawMouseIndex = Pick(latest.RawMouseIndex, baseline.RawMouseIndex, draft.RawMouseIndex),
            NdisModeIndex = affinity.NdisModeIndex,
            RssQueues = affinity.RssQueues,
            NicItrText = Pick(latest.NicItrText, baseline.NicItrText, draft.NicItrText),
        };
    }
}
