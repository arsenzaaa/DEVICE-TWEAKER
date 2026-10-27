using System.Drawing.Imaging;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;

internal static class ControlContracts
{
    private static readonly BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;

    [STAThread]
    public static int Run(string[] args)
    {
        if (args.Length != 2) throw new ArgumentException("production DLL, output directory required");
        string output = Path.GetFullPath(args[1]);
        Directory.CreateDirectory(output);
        Environment.SetEnvironmentVariable("DEVICE_TWEAKER_QA_OUTPUT", output);
        Assembly assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
        assembly.GetType("DeviceTweakerCS.QaDesktop", true)!.GetMethod("DemandAssignedDesktop", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, null);
        if (Environment.GetEnvironmentVariable("DEVICE_TWEAKER_QA_DESKTOP") is null)
            throw new InvalidOperationException("Control contracts require the private QA desktop.");
        List<object> checks = [];
        List<object> images = [];
        int failures = 0;
        void Check(string name, object expected, object observed)
        {
            bool pass = Equals(expected, observed);
            checks.Add(new { name, status = pass ? "PASS" : "FAIL", expected, observed });
            if (!pass) failures++;
        }
        void Capture(Control control, string name)
        {
            using Bitmap bitmap = new(control.Width, control.Height);
            control.DrawToBitmap(bitmap, control.ClientRectangle);
            string path = Path.Combine(output, name + ".png");
            bitmap.Save(path, ImageFormat.Png);
            images.Add(new { path, width = bitmap.Width, height = bitmap.Height, focused = control.Focused, enabled = control.Enabled,
                sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) });
        }
        Control New(string name) => (Control)Activator.CreateInstance(assembly.GetType("DeviceTweakerCS." + name, true)!, nonPublic: true)!;
        object? Call(object target, string name, params object?[] parameters) =>
            target.GetType().GetMethod(name, Hidden)!.Invoke(target, parameters);

        using Form host = new() { Name = "QA_CONTROL_HOST", ShowInTaskbar = false,
            Size = new Size(540, 220), BackColor = Color.FromArgb(12, 12, 16) };
        using Control picker = New("ThemedDropDownPicker");
        picker.Bounds = new Rectangle(20, 20, 240, 30);
        picker.BackColor = Color.FromArgb(18, 18, 22);
        picker.ForeColor = Color.FromArgb(240, 240, 240);
        var items = (List<object>)picker.GetType().GetProperty("Items")!.GetValue(picker)!;
        items.AddRange(["Enabled", "Disabled", "Windows Default"]);
        picker.GetType().GetProperty("SelectedIndex")!.SetValue(picker, 0);
        using Control toggle = New("ThemedCheckBox");
        toggle.Text = "Включено";
        toggle.Bounds = new Rectangle(20, 70, 200, 26);
        toggle.BackColor = host.BackColor;
        toggle.ForeColor = Color.White;
        using TextBox next = new() { Bounds = new Rectangle(280, 20, 170, 30) };
        host.Controls.AddRange([picker, toggle, next]);
        host.Show();
        Application.DoEvents();
        try
        {
            Check("native editors and labels preserve full glyphs at 100/125/150%", true, NativeEditorEvidence.Run(assembly, host, output));
            foreach (Keys key in new[] { Keys.Up, Keys.Down, Keys.Home, Keys.End, Keys.PageUp, Keys.PageDown })
                Check("picker routes " + key + " as input", true, Call(picker, "IsInputKey", key)!);
            Call(picker, "OnMouseDown", new MouseEventArgs(MouseButtons.Right, 1, 5, 5, 0));
            Check("right click leaves picker closed", true, picker.GetType().GetField("_popup", Hidden)!.GetValue(picker) is null);
            Call(picker, "ClosePopup");
            picker.Focus();
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Space));
            Check("Space opens picker", true, picker.GetType().GetField("_popup", Hidden)!.GetValue(picker) is Form { Visible: true });
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Down));
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Enter));
            Check("Down then Enter commits next item", 1, picker.GetType().GetProperty("SelectedIndex")!.GetValue(picker)!);
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Space));
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Escape));
            Check("Escape closes without changing selection", true, picker.GetType().GetField("_popup", Hidden)!.GetValue(picker) is null);
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Space));
            next.Focus();
            Application.DoEvents();
            Check("leaving picker closes popup", true, picker.GetType().GetField("_popup", Hidden)!.GetValue(picker) is null);
            Call(picker, "ClosePopup");
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Down));
            Check("closed picker Down selects the next item", 2, picker.GetType().GetProperty("SelectedIndex")!.GetValue(picker)!);
            Call(picker, "OnKeyDown", new KeyEventArgs(Keys.Home));
            Check("closed picker Home selects the first item", 0, picker.GetType().GetProperty("SelectedIndex")!.GetValue(picker)!);
            foreach (Control control in new[] { picker, toggle })
            {
                string name = control.GetType().Name;
                next.Focus();
                Call(control, "OnMouseLeave", EventArgs.Empty);
                Application.DoEvents();
                Check(name + " normal capture has no focus", false, control.Focused);
                Capture(control, name + "-normal");
                Call(control, "OnMouseEnter", EventArgs.Empty);
                Capture(control, name + "-hover");
                Call(control, "OnMouseLeave", EventArgs.Empty);
                control.Focus();
                // Show keyboard focus cues on the owned QA window only.
                Message cues = Message.Create(control.Handle, 0x0127, new IntPtr(0x00030002), IntPtr.Zero);
                typeof(Control).GetMethod("WndProc", Hidden)!.Invoke(control, [cues]);
                Check(name + " focus capture has focus", true, control.Focused);
                Capture(control, name + "-focus");
                control.Enabled = false;
                Capture(control, name + "-disabled");
                control.Enabled = true;
            }
            using Form main = (Form)Activator.CreateInstance(assembly.GetType("DeviceTweakerCS.MainForm", true)!)!;
            using TextBox editor = new() { Text = "0xC8" };
            main.Controls.Add(editor);
            main.Show();
            // Exercise the form's real key-preview handler without device enumeration.
            (main.GetType().GetField("_startupRefreshTimer", Hidden)!.GetValue(main) as System.Windows.Forms.Timer)?.Stop();
            main.ActiveControl = editor;
            KeyEventArgs undo = new(Keys.Control | Keys.Z);
            main.GetType().GetMethod("OnMainFormKeyDown", Hidden)!.Invoke(main, [main, undo]);
            Check("Ctrl+Z stays with the active text editor", false, undo.Handled);
            Control scroll = (Control)main.GetType().GetField("_devicesScroll", Hidden)!.GetValue(main)!;
            scroll.Visible = true;
            foreach (Keys key in new[] { Keys.Home, Keys.End, Keys.PageUp, Keys.PageDown })
            {
                KeyEventArgs editKey = new(key);
                main.GetType().GetMethod("OnMainFormKeyDown", Hidden)!.Invoke(main, [main, editKey]);
                Check(key + " stays with the active text editor", false, editKey.Handled);
            }
            Type language = assembly.GetType("DeviceTweakerCS.UiLanguage", true)!;
            MethodInfo setLanguage = language.GetMethod("Set")!;
            Type languageCode = setLanguage.GetParameters()[0].ParameterType;
            using Label pending = new();
            main.Controls.Add(pending);
            pending.Text = "pending APPLY: 1K (DWORD=1)";
            Check("pending value is translated in RU", "ожидает ПРИМЕНИТЬ: 1K (DWORD=1)", pending.Text);
            setLanguage.Invoke(null, [Enum.Parse(languageCode, "English"), false]);
            Check("RU to EN preserves the pending value", "pending APPLY: 1K (DWORD=1)", pending.Text);
            setLanguage.Invoke(null, [Enum.Parse(languageCode, "Russian"), false]);
            Check("EN to RU preserves the pending value", "ожидает ПРИМЕНИТЬ: 1K (DWORD=1)", pending.Text);
            using CheckBox stateToggle = (CheckBox)New("ThemedCheckBox");
            stateToggle.Checked = true;
            stateToggle.GetType().GetProperty("SyncCheckedStateText")!.SetValue(stateToggle, true);
            main.Controls.Add(stateToggle);
            stateToggle.Checked = false;
            Check("unchecked toggle says Disabled in RU", "Выключено", stateToggle.Text);
            // Reproduce a stale source caption during construction/localization.
            stateToggle.Text = "Enabled";
            Call(main, "LocalizeControl", stateToggle);
            Check("localization uses Checked instead of a stale toggle caption", "Выключено", stateToggle.Text);
            setLanguage.Invoke(null, [Enum.Parse(languageCode, "English"), false]);
            Check("unchecked toggle says Disabled after RU to EN", "Disabled", stateToggle.Text);
            setLanguage.Invoke(null, [Enum.Parse(languageCode, "Russian"), false]);
            Check("unchecked toggle says Disabled after EN to RU", "Выключено", stateToggle.Text);
            stateToggle.Checked = true;
            Check("checked toggle says Enabled in RU", "Включено", stateToggle.Text);
            foreach (string caption in new[] { "ПОДРОБНОСТИ", "СКРЫТЬ ПОДРОБНОСТИ" })
            {
                using Button button = (Button)Call(main, "NewDialogButton", caption, 32, 116, false)!;
                int textWidth = TextRenderer.MeasureText(caption, button.Font, Size.Empty,
                    TextFormatFlags.NoPadding | TextFormatFlags.SingleLine).Width;
                Check("dialog button fits Russian caption " + caption, true, button.Width >= textWidth + 24);
            }
            using Control info = New("InfoTextBox");
            main.Controls.Add(info);
            info.Text = "PNP ID: PCI_SENTINEL\nClass: USB\nRegistry: HKLM\\QA_SENTINEL";
            object block = Activator.CreateInstance(assembly.GetType("DeviceTweakerCS.DeviceBlock", true)!)!;
            block.GetType().GetProperty("InfoLabel")!.SetValue(block, info);
            main.GetType().GetMethod("UpdateRawMouseThrottleInfoLine", Hidden)!.Invoke(main, [block]);
            setLanguage.Invoke(null, [Enum.Parse(languageCode, "English"), false]);
            Check("throttle status update retains English source for other info rows", true,
                info.Text.Contains("Class: USB", StringComparison.Ordinal) && !info.Text.Contains("Класс", StringComparison.Ordinal));
            Check("throttle status update retains a single status line", 1,
                info.Text.Split('\n').Count(line => line.StartsWith("Raw Input Throttle:", StringComparison.OrdinalIgnoreCase)));
            main.Close();
        }
        finally { host.Close(); }
        File.WriteAllText(Path.Combine(output, "controls.json"), JsonSerializer.Serialize(new
        {
            status = failures == 0 ? "PASS" : "FAIL", checks, images,
            scope = "Real production controls on private desktop: WinForms input-key classification, popup state and selection, active-editor shortcut; event-driven state captures. Physical keyboard and real Windows DPI not exercised."
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Control contracts: {checks.Count - failures}/{checks.Count}; images {images.Count}; output {output}");
        return failures == 0 ? 0 : 1;
    }
}
