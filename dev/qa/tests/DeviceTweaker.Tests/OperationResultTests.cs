using System.Reflection;
using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class OperationResultTests
{
    [Fact]
    public void PartialResultDetailsIncludeAppliedStepsWithoutTechnicalText()
    {
        OperationReport report = new();
        report.AddSuccess("APPLIED_SENTINEL", "Value saved");
        report.AddSuccess("APPLIED_WITHOUT_MESSAGE");
        report.AddSkipped("SKIP_SENTINEL", "No eligible target");
        report.AddError("ERROR_SENTINEL", "Write failed", "Windows error 5");
        MethodInfo method = typeof(MainForm).GetMethod("BuildOperationTechnicalDetails", BindingFlags.Static | BindingFlags.NonPublic)!;
        string details = (string)method.Invoke(null, [report])!;
        Assert.Contains("APPLIED_SENTINEL", details);
        Assert.Contains("Value saved", details);
        Assert.Contains("APPLIED_WITHOUT_MESSAGE", details);
        Assert.Contains("SKIP_SENTINEL", details);
        Assert.Contains("No eligible target", details);
        Assert.Contains("ERROR_SENTINEL", details);
        Assert.Contains("Windows error 5", details);
    }

    [Fact]
    public void TechnicalDetailsKeepExplicitSkipSeparateFromWarningAndError()
    {
        OperationReport report = new();
        report.AddSkipped("SKIP_SENTINEL", "not eligible");
        report.AddWarning("WARNING_SENTINEL", "restart required");
        report.AddError("ERROR_SENTINEL", "write failed");
        MethodInfo method = typeof(MainForm).GetMethod("BuildOperationTechnicalDetails", BindingFlags.Static | BindingFlags.NonPublic)!;
        string details = (string)method.Invoke(null, [report])!;
        string skip = details.Split('\n').Single(line => line.Contains("SKIP_SENTINEL", StringComparison.Ordinal));
        Assert.DoesNotContain("WARNING", skip, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ПРЕДУПРЕЖДЕНИЕ", skip, StringComparison.OrdinalIgnoreCase);
        Assert.True(skip.Contains("SKIPPED", StringComparison.Ordinal) || skip.Contains("ПРОПУЩЕНО", StringComparison.Ordinal));
        Assert.Contains("WARNING_SENTINEL", details, StringComparison.Ordinal);
        Assert.Contains("ERROR_SENTINEL", details, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Failed", "Error", false)]
    [InlineData("AppliedWithWarnings", "Warning", false)]
    [InlineData("SkippedNoUsb", "Warning", true)]
    [InlineData("SkippedNoController", "Warning", true)]
    [InlineData("SkippedNoConfig", "Warning", true)]
    [InlineData("Applied", "Success", false)]
    public void RealImodOutcomeKeepsFailureWarningAndSkipDistinct(string outcome, string severity, bool skipped)
    {
        MethodInfo method = typeof(MainForm).GetMethod("AddImodResultToReport", BindingFlags.Static | BindingFlags.NonPublic)!;
        object value = Enum.Parse(method.GetParameters()[1].ParameterType, outcome);
        OperationReport report = new();
        method.Invoke(null, [report, value, null, "Register sequence evidence", "USB IMOD"]);
        OperationIssue issue = Assert.Single(report.Issues);
        Assert.Equal(severity, issue.Severity.ToString());
        Assert.Equal(skipped, issue.IsSkipped);
        Assert.Equal(outcome != "Failed", report.Succeeded);
    }
}
