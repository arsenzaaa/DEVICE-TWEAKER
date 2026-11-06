using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private async Task DevAuditOperationFeedbackAsync(string id)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly)
            throw new InvalidOperationException("Feedback audit requires synthetic read-only controls.");
        DeviceBlock usb = _blocks.First(b => b.ImodCheckButton is not null);
        DeviceBlock nic = _blocks.First(b => b.NicItrCheckButton is not null && b.NicItrSaveButton is not null);
        (Button Button, string Caption, string Suffix)[] cases =
        [
            (DevDescendants(usb.Group).OfType<Button>().First(b => GetSourceControlText(b) == "SET"), "SET…", "usb-set"),
            (DevDescendants(usb.Group).OfType<Button>().First(b => GetSourceControlText(b) == "DELETE"), "DELETE…", "usb-delete"),
            (usb.ImodCheckButton!, "CHECK…", "usb-check"),
            (nic.NicItrApplyButton!, "SET…", "nic-set"),
            (nic.NicItrCheckButton!, "CHECK…", "nic-check"),
            (nic.NicItrSaveButton!, "SAVE…", "nic-save"),
            (_btnApplyRef!, "APPLY…", "apply"),
            (_btnAutoRef!, "AUTO…", "auto"),
            (_btnRestoreRef!, "RESTORING…", "restore"),
        ];
        List<object> results = [];
        foreach ((Button button, string caption, string suffix) in cases)
        {
            string original = GetSourceControlText(button);
            string statusBefore = nic.NicItrStatusLabel?.Text ?? string.Empty;
            bool enabledBefore = button.Enabled;
            DevReveal(button);
            using (IDisposable operation = TryBeginDeviceOperation("FEEDBACK AUDIT")
                ?? throw new InvalidOperationException("Feedback audit could not acquire operation."))
            {
                try
                {
                    using IDisposable feedback = BeginButtonFeedback(button, caption);
                    if (button.Enabled || button.Text != UiLanguage.Text(caption))
                        throw new InvalidOperationException("Busy caption was not immediate.");
                    if (TryBeginDeviceOperation("REPEATED AUDIT") is not null)
                        throw new InvalidOperationException("Operation guard allowed duplicate work.");
                    await Task.Delay(60);
                    int textWidth = TextRenderer.MeasureText(button.Text, button.Font, Size.Empty,
                        TextFormatFlags.SingleLine).Width;
                    if (textWidth > button.Width - button.Padding.Horizontal - 8 || !DevInViewport(button))
                        throw new InvalidOperationException("Busy caption is clipped or outside the viewport: " + suffix);
                    DevCaptureWindow(id + "-" + suffix, "@main", false);
                    UiLanguageCode language = UiLanguage.Current;
                    UiLanguage.Set(language == UiLanguageCode.English ? UiLanguageCode.Russian : UiLanguageCode.English, persist: false);
                    if (button.Text != UiLanguage.Text(caption))
                        throw new InvalidOperationException("Language change lost busy feedback.");
                    UiLanguage.Set(language, persist: false);
                    throw new InvalidOperationException("Modeled operation failure");
                }
                catch (InvalidOperationException error) when (error.Message == "Modeled operation failure") { }
                if (button.Text != UiLanguage.Text(original))
                    throw new InvalidOperationException("Failure exit did not restore the button caption.");
            }
            if (button.Enabled != enabledBefore || (nic.NicItrStatusLabel?.Text ?? string.Empty) != statusBefore)
                throw new InvalidOperationException("Feedback changed enabled state or hardware result text.");
            results.Add(new { suffix, restored = true, duplicateBlocked = true, resultPreserved = true });
        }
        using Button disposed = NewCardButton(UiText.Check, UiScale(76, 24));
        using (BeginButtonFeedback(disposed, "CHECK…")) disposed.Dispose();
        if (_busyButtonCaptions.Count != 0 || _operations.IsBusy)
            throw new InvalidOperationException("Feedback scope or operation lease leaked.");
        string relative = "state/" + id + ".json";
        DevFiles.Write(DevArtifact(relative), new { cases = results,
            scope = "Production feedback helper and operation coordinator with native synthetic controls; no backup, driver or registry writes" });
        DevRecordState(relative);
    }
}
