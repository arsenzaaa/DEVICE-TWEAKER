using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private async Task DevAuditEditors(string id)
    {
        QaDesktop.DemandAssignedDesktop();
        if (!QaRun.IsReadOnly || !_testDevicesOnly) throw new InvalidOperationException("Live editor audit requires the read-only sandbox");
        DevFiles.Id(id);
        TextBox[] editors = DevDescendants(this).OfType<TextBox>()
            .Where(e => !e.Multiline && e.Visible && (e.Parent is ThemedTextBox or ThemedNumericUpDown)).ToArray();
        if (editors.Length == 0) throw new InvalidOperationException("No production editors were audited");
        List<object> cases = [];
        List<string> failures = [];
        Button focusTarget = _btnScanRef ?? throw new InvalidOperationException("Audit focus target is unavailable");
        string folder = "state/" + id;
        Directory.CreateDirectory(DevArtifact(folder));
        for (int i = 0; i < editors.Length; i++)
        {
            TextBox editor = editors[i];
            Control field = editor.Parent!;
            string family = ReferenceEquals(field, _searchFilterBox) ? "search"
                : field is ThemedNumericUpDown ? "rss-queues"
                : field.Name.StartsWith("IMOD_ROLE_", StringComparison.Ordinal) ? field.Name
                : _blocks.Any(b => ReferenceEquals(b.LimitBox, editor)) ? "msi-limit"
                : _blocks.Any(b => ReferenceEquals(b.ImodBox, editor)) ? "imod"
                : _blocks.Any(b => ReferenceEquals(b.NicItrBox, editor)) ? "nic-itr"
                : "unknown";
            if (family == "unknown") throw new InvalidOperationException("Unclassified production editor: " + field.Name);
            try { DevReveal(field); }
            catch (Exception ex) { throw new InvalidOperationException($"Editor {i}:{family} cannot fit in its viewport", ex); }
            focusTarget.Focus();
            await Task.Delay(30); // Drain the real unfocused-text refresh callbacks.
            DeviceBlock? block = _blocks.FirstOrDefault(b => b.Group.Contains(field));
            string stem = $"{folder}/{i:D3}-{family}";
            string[] paths = [stem+"-field.png",stem+"-actual.png",stem+"-reference.png",stem+"-after-focus.png"];
            using (NativeGlyphEvidence.Capture(field, DevArtifact(paths[0]))) { }
            using Bitmap actual = NativeGlyphEvidence.Capture(editor, DevArtifact(paths[1]));
            using TextBox reference = new()
            {
                Font=editor.Font, Text=editor.Text, PlaceholderText=editor.PlaceholderText,
                BorderStyle=BorderStyle.None, AutoSize=false, Multiline=false,
                BackColor=editor.BackColor, ForeColor=editor.ForeColor, Enabled=editor.Enabled,
                ReadOnly=editor.ReadOnly, TextAlign=editor.TextAlign, HideSelection=true, TabStop=false,
                Bounds=new Rectangle(8,8,editor.Width,Math.Max(editor.Height,editor.PreferredHeight)+16),
            };
            Controls.Add(reference);
            reference.BringToFront();
            using Bitmap complete = NativeGlyphEvidence.Capture(reference, DevArtifact(paths[2]));
            string expected = NativeGlyphEvidence.Signature(complete);
            bool same = NativeGlyphEvidence.Signature(actual) == expected;
            Controls.Remove(reference);
            reference.Dispose();
            bool contained = editor.Top >= 0 && editor.Bottom <= field.ClientSize.Height;
            bool focused = false;
            int start = editor.SelectionStart, length = editor.SelectionLength;
            string? focusPath = null;
            if (editor.Enabled)
            {
                focused = editor.Focus();
                editor.Select(editor.TextLength,0);
                focusPath=stem+"-focused.png";
                using (NativeGlyphEvidence.Capture(field,DevArtifact(focusPath))) { }
                editor.Select(start,length);
                focusTarget.Focus();
            }
            await Task.Delay(30);
            using Bitmap after = NativeGlyphEvidence.Capture(editor, DevArtifact(paths[3]));
            bool afterSame = NativeGlyphEvidence.Signature(after) == expected;
            bool rendered = expected.Length > 0 || (editor.TextLength == 0 && string.IsNullOrEmpty(editor.PlaceholderText));
            if (!same || !contained || !afterSame || !rendered || (editor.Enabled && !focused))
                failures.Add($"{i}:{family} complete={same} contained={contained} afterFocus={afterSame} rendered={rendered} focus={focused}");
            cases.Add(new { index=i,family,device=block?.Device.InstanceId ?? "main",editor.Text,editor.PlaceholderText,
                editor.Enabled,editor.ReadOnly,font=editor.Font.ToString(),outer=field.Bounds.ToString(),inner=editor.Bounds.ToString(),
                editor.PreferredHeight,completeGlyph=same,contained,afterFocus=afterSame,focused,rendered,paths,focusPath });
            foreach (string path in paths) DevRecordState(path);
            if (focusPath is not null) DevRecordState(focusPath);
            if (family is "msi-limit" or "nic-itr" or "search")
                DevCaptureWindow($"{id}-{i:D3}-context", "@main", false);
        }
        string report=folder+"/editors.json";
        DevFiles.Write(DevArtifact(report),new { status=failures.Count==0?"PASS":"FAIL",scope="Actual production card editors, native RGB glyphs versus taller references, clipping and focus/blur; synthetic devices only.",cases,failures });
        DevRecordState(report);
        if (failures.Count != 0) throw new InvalidOperationException("Live editor audit: "+string.Join("; ",failures));
    }
}
