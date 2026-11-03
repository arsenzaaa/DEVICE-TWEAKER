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

    internal static DeviceEditorState MergeDraft(DeviceEditorState latest, DeviceEditorState baseline, DeviceEditorState draft)
    {
        static T Pick<T>(T current, T before, T edited) => EqualityComparer<T>.Default.Equals(before, edited) ? current : edited;
        return latest with
        {
            AffinityMask = Pick(latest.AffinityMask, baseline.AffinityMask, draft.AffinityMask),
            MsiIndex = Pick(latest.MsiIndex, baseline.MsiIndex, draft.MsiIndex),
            LimitText = Pick(latest.LimitText, baseline.LimitText, draft.LimitText),
            PriorityIndex = Pick(latest.PriorityIndex, baseline.PriorityIndex, draft.PriorityIndex),
            PolicyIndex = Pick(latest.PolicyIndex, baseline.PolicyIndex, draft.PolicyIndex),
            PowerSaving = Pick(latest.PowerSaving, baseline.PowerSaving, draft.PowerSaving),
            ImodText = Pick(latest.ImodText, baseline.ImodText, draft.ImodText),
            ImodEnabled = Pick(latest.ImodEnabled, baseline.ImodEnabled, draft.ImodEnabled),
            RawMouseEnabled = Pick(latest.RawMouseEnabled, baseline.RawMouseEnabled, draft.RawMouseEnabled),
            RawMouseIndex = Pick(latest.RawMouseIndex, baseline.RawMouseIndex, draft.RawMouseIndex),
            NdisModeIndex = Pick(latest.NdisModeIndex, baseline.NdisModeIndex, draft.NdisModeIndex),
            RssQueues = Pick(latest.RssQueues, baseline.RssQueues, draft.RssQueues),
            NicItrText = Pick(latest.NicItrText, baseline.NicItrText, draft.NicItrText),
        };
    }
}
