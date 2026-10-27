namespace DeviceTweakerCS;

using DeviceTweakerQa;

public sealed partial class MainForm
{
    private const string DevBackupPath = @"C:\Example\Projects\DEVICE TWEAKER\build\Publish\DEVICE TWEAKER (SELF-CONTAINED)\Backups\DeviceTweakerBackup_20261024_204041_097_pre-auto.json";

    private void DevAuditDialog(string id)
    {
        Form window = DevActiveForm();
        List<string> defects = [];
        if (!Screen.FromControl(window).WorkingArea.Contains(window.Bounds)) defects.Add("Dialog extends outside the monitor working area.");
        List<object> controls = [];
        foreach (Control control in DevDescendants(window).Where(c => c.Visible && c is Label or Button))
        {
            bool fits = true;
            if (control is Button)
                fits = DevInViewport(control) && TextRenderer.MeasureText(control.Text, control.Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.SingleLine).Width <= control.Width - UiScale(8);
            else if (control is Label label && !label.AutoEllipsis && !string.IsNullOrWhiteSpace(label.Text))
                fits = label.GetPreferredSize(new Size(label.Width, 0)).Height <= label.Height + 1;
            controls.Add(new { control.Name, control.Text, bounds = control.Bounds.ToString(), fits });
            if (!fits) defects.Add(control.GetType().Name + ": " + control.Name + " / " + control.Text);
        }
        string relative = "state/" + id + ".layout.json";
        DevFiles.Write(DevArtifact(relative), new { window = window.Name, scale = _uiScale, controls, defects });
        DevRecordState(relative);
        if (defects.Count != 0) throw new InvalidOperationException("Dialog layout defects: " + string.Join("; ", defects));
    }

    private void DevPreviewDialog(string name)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly) throw new InvalidOperationException("Dialog preview requires synthetic read-only mode.");
        DriverEnvironmentSnapshot environment = new(new(0, "fixture"), new(null, "missing"), false,
            new("WinDefend", true, "stopped"), new("FACEIT", false, "not found"),
            new("vgk", true, "running"), new("vgc", true, "stopped"));
        switch (name)
        {
            case "backup-restored": ShowBackupFileResult(DevBackupPath, restored: true); break;
            case "backup-saved": ShowBackupFileResult(DevBackupPath, restored: false); break;
            case "system": ShowDriverEnvironmentDetails(this, environment); break;
            case "system-unknown": ShowDriverEnvironmentDetails(this, environment with { Blocklist = new(null, "unavailable"), HvciRunning = null, Faceit = new("FACEIT", null, "unavailable") }); break;
            case "imod": ShowThemedConfirm(DriverEnvironment.BuildCompactConfirmation(environment), "USB IMOD TUNING", "APPLY", "SKIP", environment); break;
            case "confirm": ShowThemedConfirm("Restore the selected backup?", "RESTORE", "RESTORE", "CANCEL"); break;
            case "backup-location": ShowAutoBackupChoiceDialog(); break;
            case "restore-empty": ShowRestoreChoiceDialog([], out _); break;
            case "restore":
            case "restore-many":
            case "restore-one":
            case "restore-delete-failure":
            case "restore-delete-exception":
                List<BackupSnapshotInfo> backups = Enumerable.Range(0, name == "restore-many" ? 9 : 2).Select(i => new BackupSnapshotInfo
                {
                    Path = i == 1 ? DevBackupPath.Replace("20261024_204041_097_pre-auto", "ORIGINAL") : DevBackupPath.Replace("097_pre-auto", $"{i:000}_pre-auto"),
                    Location = i % 2 == 0 ? "EXE" : "APPDATA", Reason = i == 1 ? "original-state" : "pre-auto",
                    IsOriginal = i == 1, CreatedAt = new DateTime(2026, 10, 24, 20, 40, 41).AddMinutes(-i),
                    LastWriteUtc = new DateTime(2026, 10, 24, 17, 40, 41, DateTimeKind.Utc).AddMinutes(-i),
                }).ToList();
                if (name == "restore-one") backups.RemoveAt(1);
                bool injectedDeleteConsumed = false;
                if (name == "restore-delete-exception") WriteLog("TEST.SCENARIO.BEGIN: dialog=backup-delete fault=BackupDelete");
                RestoreChoice choice = ShowRestoreChoiceDialog(backups, out string? selectedPath, selected =>
                {
                    if (name == "restore-delete-exception")
                    {
                        injectedDeleteConsumed = true;
                        throw new IOException("Injected file deletion failure.");
                    }
                    return name == "restore-delete-failure" || selected.IsOriginal ? 0 : 1;
                });
                string relative = "state/backup-menu-" + name + "-result.json";
                DevFiles.Write(DevArtifact(relative), new { choice = choice.ToString(), selectedPath, synthetic = true });
                DevRecordState(relative);
                if (name == "restore-delete-exception")
                {
                    if (!injectedDeleteConsumed || choice != RestoreChoice.Cancel) throw new InvalidOperationException("Delete failure preview was not consumed and dismissed.");
                    WriteLog("TEST.SCENARIO.END: status=PASS assertions=2/0 faultConsumed=True");
                }
                if (choice == RestoreChoice.DeleteBackup) ShowThemedInfo("Backup files deleted: 1");
                break;
            case "restore-reasons":
                string[] reasons = ["pre-restore", "pre-release", "pre-imod", "pre-nic-itr", "pre-nic-itr-save"];
                ShowRestoreChoiceDialog(reasons.Select((reason, index) => new BackupSnapshotInfo
                {
                    Path = DevBackupPath.Replace("pre-auto", reason), Location = "EXE", Reason = reason,
                    CreatedAt = new DateTime(2026, 10, 25, 9, 0, 0).AddMinutes(-index),
                    LastWriteUtc = new DateTime(2026, 10, 25, 6, 0, 0, DateTimeKind.Utc).AddMinutes(-index),
                }).ToArray(), out _);
                break;
            case "success":
            case "warnings":
            case "partial":
            case "failed":
            case "stress":
                OperationReport report = new() { IsDeveloperPreview = true };
                report.SetBackupPath(DevBackupPath);
                if (name != "failed") report.AddSuccess("DEVICE SETTINGS", "6 processed");
                else report.MarkNoChangesMade();
                if (name == "warnings") report.AddWarning("USB POWER", "Some power settings require a device restart.");
                if (name is "partial" or "failed") report.AddError("USB IMOD", "The register value could not be verified.", "Readback mismatch: expected 0xC8; received 0xFA0.");
                if (name == "stress") for (int i = 1; i <= 14; i++) report.AddError($"COMPONENT {i:00}", "The register value could not be verified.", $"Readback #{i}: " + new string('A', 420));
                ShowOperationResult(report, "Please reboot your PC to finish applying the changes.", name == "failed" ? "The operation stopped before any changes were made." : "Applied steps were saved. One optional step was not completed.", operationName: "APPLY");
                break;
            default: throw new InvalidDataException("Unknown dialog preview: " + name);
        }
    }
}
