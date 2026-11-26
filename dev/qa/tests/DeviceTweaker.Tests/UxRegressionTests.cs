using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class UxRegressionTests
{
    [Fact]
    public void FocusHygiene_ReleasingFocusedCardButtonDoesNotMoveFocusOntoMouseCard()
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

                DeviceInfo[] devices =
                [
                    new()
                    {
                        Name = "HID-compliant mouse (0x0)",
                        InstanceId = @"PCI\VEN_10EC&DEV_8125&SUBSYS_7D361462&REV_05\4&17112BEF&0&00E2",
                        Class = "Net",
                        RegBase = string.Empty,
                        Kind = DeviceKind.NET_NDIS,
                    },
                    new()
                    {
                        Name = "Intel(R) USB 3.20 eXtensible Host Controller - 1.20",
                        InstanceId = @"PCI\VEN_8086&DEV_7AE0&SUBSYS_7D361462&REV_11\3&11583659&0&A0",
                        Class = "USB",
                        RegBase = string.Empty,
                        Kind = DeviceKind.USB,
                        UsbRoles = "Mouse 8K",
                        UsbIsXhci = true,
                        UsbHasDevices = true,
                    },
                ];

                SetSimpleCpuTopology(form, devices);

                MethodInfo newBlock = typeof(MainForm).GetMethod("NewDeviceBlock", BindingFlags.Instance | BindingFlags.NonPublic)!;
                for (int i = 0; i < devices.Length; i++)
                {
                    newBlock.Invoke(form, [devices[i], i, null, null]);
                }

                MethodInfo layoutBlocks = typeof(MainForm).GetMethod("LayoutBlocks", BindingFlags.Instance | BindingFlags.NonPublic)!;
                layoutBlocks.Invoke(form, [true]);

                var blocks = Assert.IsType<List<DeviceBlock>>(
                    typeof(MainForm).GetField("_blocks", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form));
                DeviceBlock mouse = blocks[0];
                DeviceBlock usb = blocks[1];
                Button usbCheck = usb.ImodCheckButton!;

                var usbContainer = Assert.IsAssignableFrom<ContainerControl>(usb.Group);
                form.ActiveControl = usbCheck;
                Assert.Same(usb.Group, form.ActiveControl);
                Assert.Same(usbCheck, usbContainer.ActiveControl);

                MethodInfo releaseFocus = typeof(MainForm).GetMethod("ReleaseControlFocus", BindingFlags.Instance | BindingFlags.NonPublic)!;
                releaseFocus.Invoke(form, [usbCheck]);

                Assert.Null(form.ActiveControl);
                Assert.Null(usbContainer.ActiveControl);
                Assert.False(mouse.Group.ContainsFocus);
                Assert.False(usb.Group.ContainsFocus);

                MethodInfo beginFeedback = typeof(MainForm).GetMethod("BeginButtonFeedback", BindingFlags.Instance | BindingFlags.NonPublic)!;
                using (var _ = (IDisposable)beginFeedback.Invoke(form, [usbCheck, "CHECK…", null])!)
                {
                    Assert.False(usbCheck.Enabled);
                    Assert.Null(form.ActiveControl);
                    Assert.False(mouse.Group.ContainsFocus);
                }

                Assert.True(usbCheck.Enabled);
                Assert.Equal(UiLanguage.Text("CHECK"), usbCheck.Text);
            }
            finally
            {
                UiLanguage.Set(previousLanguage, persist: false);
            }
        });
    }

    [Fact]
    public void RawInputBackgroundMonitor_DisposesCleanlyWithoutDelay()
    {
        var samples = 0;
        var sw = Stopwatch.StartNew();
        using (var monitor = new RawInputInterop.RawInputBackgroundMonitor((_, _) => Interlocked.Increment(ref samples)))
        {
            // Dispose immediately and verify the message-only loop exits without stalling teardown.
        }
        sw.Stop();
        Assert.True(sw.ElapsedMilliseconds < 1500, $"RawInputBackgroundMonitor teardown took too long: {sw.ElapsedMilliseconds} ms");
    }

    [Fact]
    public void ImodReveal_IsIdempotentAndKeepsDetailsExpandedOnRepeatedCalls()
    {
        RunSta(() =>
        {
            using MainForm form = new();
            form.Size = new Size(1172, 875);
            form.CreateControl();
            form.PerformLayout();
            DeviceInfo usbDevice = new()
            {
                Name = "AMD USB 3.10 eXtensible Host Controller - 1.10",
                InstanceId = @"PCI\VEN_1022&DEV_15E0&SUBSYS_00001022&REV_00\4&1111111&0&0041",
                Class = "USB",
                RegBase = string.Empty,
                Kind = DeviceKind.USB,
                UsbRoles = "Mouse 8K, Keyboard 1K",
                UsbIsXhci = true,
                UsbHasDevices = true,
            };

            SetSimpleCpuTopology(form, [usbDevice]);

            MethodInfo newBlock = typeof(MainForm).GetMethod("NewDeviceBlock", BindingFlags.Instance | BindingFlags.NonPublic)!;
            newBlock.Invoke(form, [usbDevice, 0, null, null]);
            MethodInfo layoutBlocks = typeof(MainForm).GetMethod("LayoutBlocks", BindingFlags.Instance | BindingFlags.NonPublic)!;
            layoutBlocks.Invoke(form, [true]);

            var blocks = Assert.IsType<List<DeviceBlock>>(
                typeof(MainForm).GetField("_blocks", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form));
            DeviceBlock usb = blocks[0];
            Assert.False(usb.ImodDetailsExpanded);

            MethodInfo reveal = typeof(MainForm).GetMethod("RevealImodCheckResult", BindingFlags.Instance | BindingFlags.NonPublic)!;
            reveal.Invoke(form, [usb]);
            Assert.True(usb.ImodDetailsExpanded);
            Assert.True(usb.ImodReadbackRequested);

            // Subsequent SET / APPLY / AUTO calls must leave details expanded instead of collapsing them.
            reveal.Invoke(form, [usb]);
            Assert.True(usb.ImodDetailsExpanded);
            Assert.True(usb.ImodReadbackRequested);
        });
    }

    [Theory]
    [InlineData("SET", "SET…", 54, 30)]
    [InlineData("SAVE", "SAVE…", 58, 26)]
    [InlineData("CHECK", "CHECK…", 76, 30)]
    [InlineData("DELETE PROFILE", "DELETE…", 76, 30)]
    [InlineData("APPLY", "APPLY…", 80, 24)]
    [InlineData("AUTO-OPTIMIZATION", "AUTO…", 80, 24)]
    [InlineData("RESTORE", "RESTORING…", 80, 24)]
    public void RussianLocalization_CardAndActionBusyCaptionsTranslateAndFitMeasuredWidth(
        string idleSource,
        string busySource,
        int minWidth,
        int horizontalPadding)
    {
        RunSta(() =>
        {
            UiLanguageCode previous = UiLanguage.Current;
            UiLanguage.Set(UiLanguageCode.Russian, persist: false);
            try
            {
                string localizedIdle = UiLanguage.Text(idleSource);
                string localizedBusy = UiLanguage.Text(busySource);
                Assert.NotEqual(idleSource, localizedIdle);
                Assert.NotEqual(busySource, localizedBusy);

                using MainForm form = new();
                using Button button = new() { Font = form.Font, Text = localizedIdle };
                MethodInfo measure = typeof(MainForm).GetMethod("MeasureCardRowButtonWidth", BindingFlags.Instance | BindingFlags.NonPublic)!;
                int measuredWidth = (int)measure.Invoke(form, [button, minWidth, horizontalPadding, idleSource, busySource])!;

                int busyTextWidth = TextRenderer.MeasureText(
                    localizedBusy,
                    button.Font,
                    Size.Empty,
                    TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
                Assert.True(
                    measuredWidth >= busyTextWidth + 16,
                    $"Button width {measuredWidth} is too narrow for busy caption '{localizedBusy}' ({busyTextWidth}px).");
            }
            finally
            {
                UiLanguage.Set(previous, persist: false);
            }
        });
    }

    private static void SetSimpleCpuTopology(MainForm form, DeviceInfo[] devices)
    {
        List<CpuLpInfo> lps = [];
        Dictionary<int, int> map = [];
        for (int lp = 0; lp < 8; lp++)
        {
            lps.Add(new CpuLpInfo(0, lp, lp / 2, 0, 0, 0, lp));
            map[lp] = 0;
        }

        CpuInfo info = new()
        {
            Topology = new CpuTopology(lps),
            CcdMap = map,
            CcxMap = map,
            CcdL3CacheBytes = new Dictionary<int, long>(),
        };

        typeof(MainForm).GetField("_cpuInfo", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, info);
        typeof(MainForm).GetField("_maxLogical", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(form, 8);

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
