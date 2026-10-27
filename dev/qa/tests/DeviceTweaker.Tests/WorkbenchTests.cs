using System.Drawing;
using System.Text.Json;
using DeviceTweakerQa;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class WorkbenchTests : IDisposable
{
    private readonly string _root = Path.Combine(QaOutput(), "workbench-unit-" + Guid.NewGuid().ToString("N"));
    private static string QaOutput() => DeviceTweakerCS.QaRun.OutputDirectory;
    public void Dispose() { } // Evidence is retained in the QA run directory.
    private string Folder(string name) { string path = Path.Combine(_root, name); Directory.CreateDirectory(path); return path; }
    private string Fixture(string name, bool changed = false)
    {
        string root = Folder(name);
        DevFiles.Write(Path.Combine(root, "scenario.json"), new DevScenario { Id = "fixture", Steps = [new() { Id = "screen", Action = "capture" }] });
        using Bitmap image = new(16, 16); using (Graphics g = Graphics.FromImage(image)) g.Clear(Color.Black);
        if (changed) image.SetPixel(7, 9, Color.White);
        image.Save(Path.Combine(root, "screen.png")); File.WriteAllText(Path.Combine(root, "tree.json"), "{}");
        DevRun run = new() { Status = "PASS", ReadOnly = true, ScenarioId = "fixture", Environment = "unit-fixture",
            ScenarioSha256 = DevFiles.Hash(Path.Combine(root, "scenario.json")), Steps = [new("screen", "capture", "PASS", 10, "")],
            Captures = [new("screen", "screen.png", DevFiles.Hash(Path.Combine(root, "screen.png")), 16, 16, 5, "tree.json", DevFiles.Hash(Path.Combine(root, "tree.json")))] };
        DevFiles.Write(Path.Combine(root, "session.json"), run);
        DevFiles.Write(Path.Combine(root, "process-1.json"), new { readOnly = true, build = new { hasTestPanel = true }, desktop = new { isolated = true } });
        DevFiles.Write(Path.Combine(root, "execution.json"), new DevExecution(DevFiles.Hash(Path.Combine(root, "session.json")), "process-1.json", 1, 0));
        DevFiles.Write(Path.Combine(root, "development.json"), new { functionalStatus = "PASS" }); return root;
    }

    [Fact]
    public void ExactComparisonDetectsOneChangedPixelAndItsBounds()
    {
        using Bitmap a = new(16, 16), b = new(16, 16); b.SetPixel(7, 9, Color.White);
        PixelDifference d = VisualComparison.Compare(a, b, Path.Combine(Folder("pixels"), "diff.png"));
        Assert.Equal("CHANGED", d.Status); Assert.Equal(1L, d.ChangedPixels); Assert.Equal((7, 9, 1, 1), (d.Left, d.Top, d.Width, d.Height));
    }
    [Fact]
    public void AlphaChangesAreNotIgnored()
    {
        using Bitmap a = new(8, 8), b = new(8, 8); b.SetPixel(0, 0, Color.Black);
        Assert.Equal("CHANGED", VisualComparison.Compare(a, b).Status);
    }
    [Fact]
    public void EqualImagesMatchAndSizeChangesCannotPass()
    {
        using Bitmap a = new(8, 8), b = new(8, 8), larger = new(9, 8);
        Assert.Equal("MATCH", VisualComparison.Compare(a, b).Status);
        Assert.Equal("SIZE_CHANGED", VisualComparison.Compare(a, larger).Status);
    }
    [Fact]
    public void ChangedRunNeedsReviewAndDoesNotReplaceBaseline()
    {
        string before = Fixture("before"), after = Fixture("after", true);
        string hash = DevFiles.Hash(Path.Combine(before, "screen.png"));
        VisualComparison.CompareRuns(before, after, Path.Combine(after, "comparison"));
        using JsonDocument result = JsonDocument.Parse(File.ReadAllText(Path.Combine(after, "comparison/comparison.json")));
        Assert.Equal("REVIEW_REQUIRED", result.RootElement.GetProperty("status").GetString());
        Assert.Equal(hash, DevFiles.Hash(Path.Combine(before, "screen.png")));
        Assert.True(File.Exists(Path.Combine(after, "comparison/index.html")));
    }
    [Fact]
    public void NoOpSameDirectoryComparisonIsRejected()
    {
        string before = Fixture("same");
        Assert.Throws<InvalidDataException>(() => VisualComparison.CompareRuns(before, before, Path.Combine(before, "comparison")));
    }
    [Fact]
    public void EnvironmentMismatchIsNotAnImageRegressionOrAPass()
    {
        string before = Fixture("env-before"), after = Fixture("env-after");
        var run = DevFiles.Read<DevRun>(Path.Combine(after, "session.json")); run.Environment = "different dpi";
        DevFiles.Write(Path.Combine(after, "session.json"), run);
        Assert.Throws<InvalidDataException>(() => VisualComparison.CompareRuns(before, after, Path.Combine(after, "comparison")));
    }
    [Theory]
    [InlineData("screen.png")] [InlineData("tree.json")]
    public void AlteredOrMissingEvidenceCannotPass(string artifact)
    {
        string root = Fixture("alter-" + artifact.Replace('.', '-'));
        File.AppendAllText(Path.Combine(root, artifact), "tampered");
        Assert.Throws<InvalidDataException>(() => DevFiles.Verify(root, DevFiles.Read<DevRun>(Path.Combine(root, "session.json"))));
        File.Delete(Path.Combine(root, artifact));
        Assert.ThrowsAny<IOException>(() => DevFiles.Verify(root, DevFiles.Read<DevRun>(Path.Combine(root, "session.json"))));
    }
    [Fact]
    public void DroppedStepCannotBeHiddenBySuccessfulStatus()
    {
        string root = Fixture("dropped"); var run = DevFiles.Read<DevRun>(Path.Combine(root, "session.json")); run.Steps.Clear();
        Assert.Throws<InvalidDataException>(() => DevFiles.Verify(root, run));
    }
    [Theory]
    [InlineData("../outside.png")] [InlineData("C:\\Windows\\test.png")]
    public void ArtifactPathCannotEscapeTheSandbox(string path)
        => Assert.Throws<InvalidDataException>(() => DevFiles.Inside(Folder("paths"), path));
    [Fact]
    public void UnknownActionsAndDuplicateStepsAreRejected()
    {
        DevScenario scene = new() { Id = "bad", Steps = [new() { Id = "s", Action = "shell" }] };
        Assert.Throws<InvalidDataException>(scene.Validate);
        scene.Steps = [new() { Id = "s", Action = "capture" }, new() { Id = "s", Action = "capture" }];
        Assert.Throws<InvalidDataException>(scene.Validate);
    }
    [Fact]
    public void UnknownJsonPropertiesAreRejectedInsteadOfSilentlySkipped()
        => Assert.Throws<JsonException>(() => JsonSerializer.Deserialize<DevScenario>("{\"id\":\"x\",\"aply\":true}", DevFiles.Json));
    [Fact]
    public void PendingOrIncompleteVisualReviewCannotApproveASession()
    {
        string root = Fixture("review-pending"), input = Path.Combine(root, "input.json");
        DevelopmentReview.Template(root, input);
        DevReview review = DevFiles.Read<DevReview>(input); review.Reviewer = "unit fixture"; review.Summary = "Pending verdict must fail";
        DevFiles.Write(input, review);
        Assert.Throws<InvalidDataException>(() => DevelopmentReview.Complete(root, input));
        Assert.False(File.Exists(Path.Combine(root, "visual-review.json")));
    }
    [Fact]
    public void DefectReviewIsRejectedAndApprovedReviewRemainsBoundToExactImages()
    {
        string root = Fixture("review-defect"), input = Path.Combine(root, "input.json"); DevelopmentReview.Template(root, input);
        DevReview review = DevFiles.Read<DevReview>(input); review.Reviewer = "unit fixture"; review.Summary = "Observed a defect";
        review.Captures = review.Captures.Select(c => c with { Verdict = "defect", Note = "Deliberate unit fixture verdict" }).ToList();
        DevFiles.Write(input, review); Assert.Equal("REJECTED", DevelopmentReview.Complete(root, input));
        review.Captures = review.Captures.Select(c => c with { Verdict = "ok", Note = "Fixture rechecked" }).ToList();
        DevFiles.Write(input, review); Assert.Equal("REVIEWED", DevelopmentReview.Complete(root, input));
        review.SessionSha256 = "wrong"; DevFiles.Write(input, review);
        Assert.Throws<InvalidDataException>(() => DevelopmentReview.Complete(root, input));
    }
    [Fact]
    public void CrashedExecutionCannotBeUsedAsABaselineEvenWithPassingScenario()
    {
        string before = Fixture("crash-before"), after = Fixture("crash-after");
        File.WriteAllText(Path.Combine(before, "crash_fixture.txt"), "Explicit unit fixture, not a real crash");
        Assert.Throws<InvalidDataException>(() => VisualComparison.CompareRuns(before, after, Path.Combine(after, "comparison")));
    }
}
