namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    // Per-role IMOD boxes (Mouse, Keyboard, Audio, Gamepad) under the IMOD row of a USB
    // controller card. They mirror the IMOD text field in both directions.
    private sealed class ImodRoleEditor
    {
        private readonly MainForm _form;
        private readonly IReadOnlyList<string> _roles;
        private readonly ThemedTextBox _text;
        private readonly CheckBox _enabled;
        private readonly ThemedDropDownPicker _mode;
        private bool _suppress;

        public Dictionary<string, Label> Labels { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, ThemedTextBox> Boxes { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Panel Panel { get; }

        public ImodRoleEditor(MainForm form, IReadOnlyList<string> roles, ThemedTextBox text, CheckBox enabled, ThemedDropDownPicker mode)
        {
            _form = form;
            _roles = roles;
            _text = text;
            _enabled = enabled;
            _mode = mode;
            Panel = new Panel
            {
                BackColor = form._bgGroup,
                Visible = roles.Count > 0,
            };

            if (_roles.Count > 0)
            {
                int editorLabelWidth = Math.Max(
                    _form.UiScale(92),
                    _roles.Max(role => TextRenderer.MeasureText(
                        UiLanguage.Text(role + ":"),
                        _form._baseFont,
                        Size.Empty,
                        TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width) + _form.UiScale(8));
                int editorBoxWidth = _form.UiScale(78);
                int editorColumnGap = _form.UiScale(18);
                int editorRowHeight = _form.UiScale(28);
                // Keep the editor usable on narrower windows: one or two roles fit on one row,
                // while three or more use a stable two-column grid. This prevents the third
                // value box from being clipped when Russian labels are wider.
                bool twoColumnGrid = _roles.Count >= 3;
                int editorX = 0;
                int editorY = 0;
                for (int i = 0; i < _roles.Count; i++)
                {
                    if (twoColumnGrid && i > 0 && i % 2 == 0)
                    {
                        editorX = 0;
                        editorY += editorRowHeight;
                    }

                    string role = _roles[i];
                    Label roleLabel = new()
                    {
                        Text = role + ":",
                        AutoSize = false,
                        Size = new Size(editorLabelWidth, _form.UiScale(22)),
                        Location = new Point(editorX, editorY + _form.UiScale(3)),
                        ForeColor = _form._fgMain,
                    };

                    ThemedTextBox roleBox = new()
                    {
                        Name = "IMOD_ROLE_" + role.ToUpperInvariant(),
                        Size = new Size(editorBoxWidth, _form.UiScale(24)),
                        Location = new Point(roleLabel.Right + _form.UiScale(4), editorY),
                        BackColor = Color.FromArgb(18, 18, 22),
                        ForeColor = _form._fgMain,
                        Text = FormatImodValue(GetImodDeviceEditorDefaultValue(role)),
                        TextAlign = HorizontalAlignment.Left,
                        Font = _form._blockFont,
                    };
                    _form.StyleDarkTextBox(roleBox);
                    roleBox.TextChanged += (_, _) => UpdateTextFromBoxes();
                    roleBox.Inner.Leave += (_, _) => Normalize(roleBox);
                    roleBox.Inner.KeyDown += (_, e) =>
                    {
                        if (e.KeyCode == Keys.Enter)
                        {
                            Normalize(roleBox);
                            e.SuppressKeyPress = true;
                        }
                    };

                    Labels[role] = roleLabel;
                    Boxes[role] = roleBox;
                    Panel.Controls.AddRange([roleLabel, roleBox]);
                    editorX = roleBox.Right + editorColumnGap;
                }
            
            }

            _text.TextChanged += (_, _) => SyncFromText();
            SyncFromText();
        }

        private void SyncFromText()
        {
            if (Boxes.Count == 0 || _suppress)
            {
                return;
            }

            bool hasRoleValues = TryParseImodRoleIntervals(_text.Text ?? string.Empty, out Dictionary<string, uint> roleValues)
                && roleValues.Count > 0;
            uint? singleValue = null;
            string imodText = _text.Text ?? string.Empty;
            if (!hasRoleValues
                && !string.IsNullOrWhiteSpace(imodText)
                && TryParseImodInterval(imodText, ImodDefaultInterval, out uint parsedSingle))
            {
                singleValue = parsedSingle;
            }

            _suppress = true;
            try
            {
                foreach (KeyValuePair<string, ThemedTextBox> pair in Boxes)
                {
                    uint value = hasRoleValues && roleValues.TryGetValue(pair.Key, out uint roleValue)
                        ? roleValue
                        : singleValue ?? GetImodDeviceEditorDefaultValue(pair.Key);
                    pair.Value.Text = FormatImodValue(value);
                    pair.Value.ForeColor = _form._fgMain;
                }
            }
            finally
            {
                _suppress = false;
            }
        
        }

        private void UpdateTextFromBoxes()
        {
            if (_suppress || Boxes.Count == 0)
            {
                return;
            }

            Dictionary<string, uint> roleValues = new(StringComparer.OrdinalIgnoreCase);
            bool valid = true;
            foreach (string role in _roles)
            {
                if (!Boxes.TryGetValue(role, out ThemedTextBox? box))
                {
                    continue;
                }

                string text = box.Text.Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                if (!TryParseImodInterval(text, ImodDefaultInterval, out uint value))
                {
                    box.ForeColor = Color.FromArgb(255, 110, 110);
                    valid = false;
                    continue;
                }

                box.ForeColor = _form._fgMain;
                roleValues[role] = value;
            }

            if (!valid || roleValues.Count == 0)
            {
                return;
            }

            _suppress = true;
            try
            {
                _text.Text = FormatImodRoleIntervals(roleValues);
                _enabled.Checked = true;
                _mode.SelectedItem = ImodModeRoles;
            }
            finally
            {
                _suppress = false;
            }
        
        }

        private void Normalize(ThemedTextBox box)
        {
            if (TryParseImodInterval(box.Text.Trim(), ImodDefaultInterval, out uint value))
            {
                _suppress = true;
                try
                {
                    box.Text = FormatImodValue(value);
                    box.ForeColor = _form._fgMain;
                }
                finally
                {
                    _suppress = false;
                }

                UpdateTextFromBoxes();
            }
        
        }
    }
}
