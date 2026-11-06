namespace DeviceTweakerCS;

// Presentation only: keep stored reason codes and backup file names unchanged.
internal static class BackupReasonLabel
{
    internal static string Get(string? reason) => UiLanguage.Text(reason?.Trim().ToLowerInvariant() switch
    {
        "pre-auto" => "Before auto-optimization",
        "pre-apply" => "Before applying settings",
        "pre-imod" => "Before USB IMOD changes",
        "pre-nic-itr" => "Before NIC ITR changes",
        "pre-nic-itr-save" => "Before saving NIC ITR profile",
        "pre-restore" => "Before restoring a backup",
        "pre-reset-tweaks" => "Before resetting tweaks",
        "original-state" => "Original device settings",
        _ => "Device settings backup",
    });
}
