namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    // This acknowledgement belongs to the current UI session, not to a hardware
    // profile. A successful load or a previous acknowledgement proves no compatibility.
    private bool _driverUseConfirmed;

    private DriverEnvironmentSnapshot CaptureDriverEnvironment()
    {
        bool? hvci = TryQueryImodCodeIntegrityOptions(out uint options)
            ? (options & ImodCiOptionHvciEnabled) != 0 : null;
        DriverEnvironmentSnapshot snapshot = DriverEnvironment.Capture(hvci);
        WriteLog("DRIVER.PREFLIGHT: " + FlattenLogText(DriverEnvironment.BuildConfirmation(snapshot)));
        return snapshot;
    }

    private bool ConfirmDriverUse()
    {
        if (_driverUseConfirmed) return true;
        if (QaRun.TryBlock("DRIVER CONSENT")) return false;
        if (InvokeRequired) throw new InvalidOperationException("Driver confirmation must run on the UI thread.");
        _driverUseConfirmed = ShowDriverUsePrompt(CaptureDriverEnvironment());
        WriteLog($"DRIVER.CONSENT: accepted={_driverUseConfirmed} scope=current-session");
        return _driverUseConfirmed;
    }

    private bool ShowDriverUsePrompt(DriverEnvironmentSnapshot snapshot)
        => ShowThemedConfirm("This action requires DTIMOD.\n\n"
            + DriverEnvironment.BriefCompatibilityNotice + "\n\n" + DriverEnvironment.BriefLifetimeNotice, "DTIMOD DRIVER",
            "CONTINUE", "CANCEL", snapshot);
}
