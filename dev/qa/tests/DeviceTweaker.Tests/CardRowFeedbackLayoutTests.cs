using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class CardRowFeedbackLayoutTests
{
    private sealed record ButtonAuditCase(string Label, string Caption, Func<DeviceBlock, Button> SelectButton, Func<DeviceBlock, Control> SelectEditor);

    [Fact]
    public void RussianCardButtonFeedbackRelayoutDoesNotExpandInputOrClipButtons()
    {
        RunSta(() =>
        {
            UiLanguageCode previousLanguage = UiLanguage.Current;
            UiLanguage.Set(UiLanguageCode.Russian, persist: false);
            using MainForm form = new();
            try
            {
                form.Size = new Size(1172, 875);
                form.CreateControl();
                form.PerformLayout();

                Set16ThreadCpuTopology(form);
                List<DeviceInfo> devices =
                [
                    new()
                    {
                        Name = "Intel(R) USB 3.20 eXtensible Host Controller - 1.20",
                        InstanceId = @"PCI\VEN_8086&DEV_7AE0&SUBSYS_7D361462&REV_11\3&11583659&0&A0",
                        Class = "USB",
                        RegBase = string.Empty,
                        Kind = DeviceKind.USB,
                        UsbRoles = "Mouse 8K, Keyboard 8K, Microphone, Audio",
                        UsbIsXhci = true,
                        UsbHasDevices = true,
                    },
                    new()
                    {
                        Name = "Realtek Gaming 2.5GbE Family Controller",
                        InstanceId = @"PCI\VEN_10EC&DEV_8125&SUBSYS_7D361462&REV_05\4&17112BEF&0&00E2",
                        Class = "Net",
                        RegBase = string.Empty,
                        Kind = DeviceKind.NET_NDIS,
                    },
                ];

                MethodInfo normalize = typeof(MainForm).GetMethod("NormalizeInstanceId", BindingFlags.Static | BindingFlags.NonPublic)!;
                var powerCache = (Dictionary<string, bool>)typeof(MainForm).GetField("_displayPowerCache", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
                var rssCache = (Dictionary<string, NdisRssRuntimeState>)typeof(MainForm).GetField("_ndisRssRuntimeCache", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
                foreach (DeviceInfo device in devices)
                {
                    string norm = (string)normalize.Invoke(null, [device.InstanceId])!;
                    powerCache[norm] = true;
                    if (device.Kind == DeviceKind.NET_NDIS)
                    {
                        rssCache[norm] = new NdisRssRuntimeState(true, true, true, 0, 0, 0, 7, 4, 4, "Ethernet", device.Name, "1", string.Empty);
                    }
                }

                MethodInfo newBlock = typeof(MainForm).GetMethod("NewDeviceBlock", BindingFlags.Instance | BindingFlags.NonPublic)!;
                for (int i = 0; i < devices.Count; i++)
                {
                    newBlock.Invoke(form, [devices[i], i, null, null]);
                }

                MethodInfo layoutBlocks = typeof(MainForm).GetMethod("LayoutBlocks", BindingFlags.Instance | BindingFlags.NonPublic)!;
                layoutBlocks.Invoke(form, [true]);

                var blocksObj = typeof(MainForm).GetField("_blocks", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form);
                var blocks = Assert.IsType<List<DeviceBlock>>(blocksObj);
                DeviceBlock usb = blocks[0];
                DeviceBlock nic = blocks[1];

                Button usbSet = usb.SettingsPanel!.Controls.OfType<Button>()
                    .First(b => string.Equals(GetSourceText(form, b), "SET", StringComparison.Ordinal));
                Button usbDelete = usb.SettingsPanel.Controls.OfType<Button>()
                    .First(b => string.Equals(GetSourceText(form, b), "DELETE PROFILE", StringComparison.Ordinal));

                ButtonAuditCase[] cases =
                [
                    new("usb-check", "CHECK…", b => b.ImodCheckButton!, b => b.ImodBox.Parent!),
                    new("usb-delete", "DELETE…", _ => usbDelete, b => b.ImodBox.Parent!),
                    new("usb-set", "SET…", _ => usbSet, b => b.ImodBox.Parent!),
                    new("nic-check", "CHECK…", b => b.NicItrCheckButton!, b => b.NicItrBox!.Parent!),
                    new("nic-save", "SAVE…", b => b.NicItrSaveButton!, b => b.NicItrBox!.Parent!),
                    new("nic-set", "SET…", b => b.NicItrApplyButton!, b => b.NicItrBox!.Parent!),
                ];

                MethodInfo beginFeedback = typeof(MainForm).GetMethod("BeginButtonFeedback", BindingFlags.Instance | BindingFlags.NonPublic)!;

                foreach (ButtonAuditCase item in cases)
                {
                    DeviceBlock block = item.Label.StartsWith("usb-", StringComparison.Ordinal) ? usb : nic;
                    Button button = item.SelectButton(block);
                    Control editor = item.SelectEditor(block);
                    Panel settingsPanel = block.SettingsPanel!;

                    Rectangle buttonBefore = button.Bounds;
                    Rectangle editorBefore = editor.Bounds;
                    string sourceBefore = GetSourceText(form, button);

                    using (var feedback = (IDisposable)beginFeedback.Invoke(form, [button, item.Caption, null])!)
                    {
                        block.RelayoutAction!.Invoke();
                        Assert.True(
                            editor.Bounds == editorBefore,
                            $"{item.Label} during feedback: expected editor bounds {editorBefore}, actual {editor.Bounds}");
                        Assert.True(
                            button.Bounds == buttonBefore,
                            $"{item.Label} during feedback: expected button bounds {buttonBefore}, actual {button.Bounds}");
                    }

                    Assert.Equal(sourceBefore, GetSourceText(form, button));
                    Assert.True(
                        editor.Bounds == editorBefore,
                        $"{item.Label} after feedback dispose: expected editor bounds {editorBefore}, actual {editor.Bounds}");
                    Assert.True(
                        button.Bounds == buttonBefore,
                        $"{item.Label} after feedback dispose: expected button bounds {buttonBefore}, actual {button.Bounds}");
                    Assert.True(
                        button.Right <= settingsPanel.ClientSize.Width,
                        $"{item.Label} clipped on right: button.Right={button.Right} > settingsWidth={settingsPanel.ClientSize.Width}");
                }

                // Also verify switching UI language mid-operation while CHECK feedback
                // is active settles cleanly to Russian idle geometry without clipping.
                Rectangle expectedRuImodBounds = usb.ImodBox.Parent!.Bounds;
                Rectangle expectedRuCheckBounds = usb.ImodCheckButton!.Bounds;
                UiLanguage.Set(UiLanguageCode.English, persist: false);
                using (var feedback = (IDisposable)beginFeedback.Invoke(form, [usb.ImodCheckButton, "CHECK…", null])!)
                {
                    UiLanguage.Set(UiLanguageCode.Russian, persist: false);
                    usb.RelayoutAction!.Invoke();
                }

                Assert.Equal(expectedRuImodBounds, usb.ImodBox.Parent!.Bounds);
                Assert.Equal(expectedRuCheckBounds, usb.ImodCheckButton!.Bounds);
                Assert.Equal("CHECK", GetSourceText(form, usb.ImodCheckButton));
                Assert.Equal(UiLanguage.Text("CHECK"), usb.ImodCheckButton.Text);
                Assert.True(usb.ImodCheckButton.Right <= usb.SettingsPanel!.ClientSize.Width);

                // Applying or checking IMOD must expand the details table cleanly
                // even when invoked programmatically (e.g. after AUTO or SET).
                Assert.False(usb.ImodDetailsExpanded);
                MethodInfo revealImod = typeof(MainForm).GetMethod("RevealImodCheckResult", BindingFlags.Instance | BindingFlags.NonPublic)!;
                revealImod.Invoke(form, [usb]);
                Assert.True(usb.ImodDetailsExpanded);
                Assert.True(usb.ImodReadbackRequested);
                Assert.Equal("HIDE", GetSourceText(form, usb.ImodDetailsButton!));
                Assert.Equal(UiLanguage.Text("HIDE"), usb.ImodDetailsButton!.Text);
            }
            finally
            {
                UiLanguage.Set(previousLanguage, persist: false);
            }
        });
    }

    private static string GetSourceText(MainForm form, Control control)
    {
        MethodInfo method = typeof(MainForm).GetMethod("GetSourceControlText", BindingFlags.Instance | BindingFlags.NonPublic)!;
        return (string)method.Invoke(form, [control])!;
    }

    private static void Set16ThreadCpuTopology(MainForm form)
    {
        List<CpuLpInfo> lps = [];
        Dictionary<int, int> ccd = [];
        Dictionary<int, int> ccx = [];
        for (int lp = 0; lp < 16; lp++)
        {
            lps.Add(new CpuLpInfo(0, lp, lp / 2, 0, 0, 0, lp));
            ccd[lp] = 0;
            ccx[lp] = 0;
        }

        CpuInfo info = new()
        {
            Topology = new CpuTopology(lps),
            CcdMap = ccd,
            CcxMap = ccx,
            CcdL3CacheBytes = new Dictionary<int, long>(),
        };

        typeof(MainForm).GetField("_cpuInfo", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, info);
        typeof(MainForm).GetField("_maxLogical", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, 16);
    }

    private static void RunSta(Action action)
    {
        Exception? failure = null;
        Thread thread = new(() =>
        {
            try { action(); }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(15000), "STA test thread timed out.");
        if (failure is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
