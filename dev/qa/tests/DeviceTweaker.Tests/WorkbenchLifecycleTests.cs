using System.Reflection;
using System.Windows.Forms;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class WorkbenchLifecycleTests
{
    private sealed class FontConsumer : Control
    {
        public bool FontObservedDuringDispose;
        protected override void Dispose(bool disposing)
        {
            if (disposing) { _ = Font.Height; FontObservedDuringDispose = true; }
            base.Dispose(disposing);
        }
    }
    [Fact]
    public void SharedFontsRemainAliveUntilChildControlsAreDisposed()
    {
        Exception? failure = null;
        Thread thread = new(() =>
        {
            try
            {
                MainForm form = new();
                var font = (System.Drawing.Font)typeof(MainForm).GetField("_blockFont", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
                FontConsumer child = new() { Font = font }; form.Controls.Add(child);
                form.Dispose();
                Assert.True(child.FontObservedDuringDispose);
                Assert.Throws<ArgumentException>(() => _ = font.Height);
            }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start();
        Assert.True(thread.Join(10000), "Hidden form disposal did not finish");
        if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
