using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class RawPollingRoleTests
{
    [Theory]
    [InlineData("Audio, Keyboard 8K, Microphone, Mouse 1K, Mouse 8K", "Mouse 1K, Keyboard 8K, Microphone, Audio")]
    [InlineData("Mouse 8K, Mouse 1K, Mouse 8K", "Mouse 1K")]
    public void ConfirmedRateDoesNotDuplicateRoleSummary(string input, string expected)
    {
        var overrides = new Dictionary<string, string> { ["Mouse"] = "1K" };
        string actual = MainForm.ApplyUsbRolePollingOverrides(input, overrides);
        Assert.Equal(expected, actual);
        Assert.Equal(actual, MainForm.ApplyUsbRolePollingOverrides(actual, overrides));
    }

    [Fact]
    public void DistinctRatesRemainWhenNoMeasurementReplacesThem()
    {
        Assert.Equal("Mouse 1K, Mouse 8K, Keyboard 8K", MainForm.ApplyUsbRolePollingOverrides(
            "Mouse 1K, Mouse 8K, Keyboard 8K", new Dictionary<string, string>()));
    }

    [Theory]
    [InlineData(1.0, 1000.0)]
    [InlineData(0.125, 8000.0)]
    public void StableIntervals_ProduceMeasuredRateDespitePausesAndBatchArtifacts(double interval, double expected)
    {
        double[] samples = Enumerable.Repeat(interval, 100)
            .Concat(new[] { 0.001, 0.002, 100.0, 500.0, double.NaN, double.PositiveInfinity }).ToArray();
        Assert.True(MainForm.TryEstimateRawPollingHz(samples, out double measured));
        Assert.Equal(expected, measured, 6);
    }

    [Fact]
    public void SparseKeyboardEvents_DoNotEstablishHighPollingRate()
    {
        double[] samples = Enumerable.Repeat(100.0, 100).ToArray();
        Assert.False(MainForm.TryEstimateRawPollingHz(samples, out _));
        Assert.True(MainForm.TryEstimateLivePollingHz(samples, out double activity));
        Assert.Equal(10.0, activity);
    }

    [Fact]
    public void TooFewIntervals_DoNotProduceAnEstimate()
    {
        Assert.False(MainForm.TryEstimateRawPollingHz(new[] { 1.0, 1.0 }, out _));
        Assert.False(MainForm.TryEstimateLivePollingHz(new[] { 1.0, 1.0 }, out _));
    }

    [Fact]
    public void ConfirmedMouseRateReplacesEndpointMaximum()
    {
        string roles = MainForm.ApplyUsbRolePollingOverrides(
            "Audio, Keyboard 8K, Microphone, Mouse 8K",
            new Dictionary<string, string> { ["Mouse"] = "1K" });

        Assert.Equal("Mouse 1K, Keyboard 8K, Microphone, Audio", roles);
    }

    [Fact]
    public void KeyboardActivityDoesNotLowerEndpointRate()
    {
        string roles = MainForm.ApplyUsbRolePollingOverrides(
            "Keyboard 8K, Mouse 8K",
            new Dictionary<string, string> { ["Keyboard"] = "1K", ["Mouse"] = "1K" });

        Assert.Equal("Mouse 1K, Keyboard 8K", roles);
    }

    [Theory]
    [InlineData("Audio, Keyboard 8K, Microphone, Mouse 1K", "Mouse 1K, Keyboard 8K, Microphone, Audio")]
    [InlineData("Webcam, Audio DAC, Gamepad, Microphone", "Gamepad, Microphone, Audio DAC, Webcam")]
    [InlineData("Audio, mouse 1K, Keyboard 8K, Mouse 1K", "mouse 1K, Keyboard 8K, Audio")]
    [InlineData("Webcam, Mousepad, Mouse 8K, Other", "Mouse 8K, Webcam, Mousepad, Other")]
    public void RoleLabelsHaveTheSameOrderBeforeAndAfterPollingConfirmation(string input, string expected)
    {
        var noMeasurement = new Dictionary<string, string>();
        string actual = MainForm.ApplyUsbRolePollingOverrides(input, noMeasurement);
        Assert.Equal(expected, actual);
        Assert.Equal(actual, MainForm.ApplyUsbRolePollingOverrides(actual, noMeasurement));
    }
}
