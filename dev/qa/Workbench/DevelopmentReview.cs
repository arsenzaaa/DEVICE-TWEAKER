using System.Text.Json;

namespace DeviceTweakerQa;

internal sealed class DevReview
{
    public string SessionSha256 { get; set; } = "";
    public string Reviewer { get; set; } = "";
    public string Summary { get; set; } = "";
    public List<DevImageReview> Captures { get; set; } = [];
}
internal sealed record DevImageReview(string Id, string Sha256, string Verdict, string Note);
internal static class DevelopmentReview
{
    public static void Template(string root, string target)
    {
        DevRun run = DevFiles.Read<DevRun>(Path.Combine(root, "session.json")); DevFiles.Verify(root, run);
        DevFiles.Write(target, new DevReview { SessionSha256 = DevFiles.Hash(Path.Combine(root, "session.json")),
            Captures = run.Captures.Select(c => new DevImageReview(c.Id, c.Sha256, "PENDING", "")).ToList() });
    }
    public static string Complete(string root, string reviewFile)
    {
        string session = Path.Combine(root, "session.json"); DevRun run = DevFiles.Read<DevRun>(session); DevFiles.Verify(root, run);
        DevFiles.VerifyExecution(root);
        DevReview review = DevFiles.Read<DevReview>(reviewFile);
        if (review.SessionSha256 != DevFiles.Hash(session) || string.IsNullOrWhiteSpace(review.Reviewer) || string.IsNullOrWhiteSpace(review.Summary))
            throw new InvalidDataException("Review must identify the exact session, reviewer and findings");
        if (review.Captures.Count != run.Captures.Count || review.Captures.Select(c => c.Id).Distinct().Count() != review.Captures.Count)
            throw new InvalidDataException("Review must cover every checkpoint / transition frame exactly once");
        foreach (DevCapture capture in run.Captures)
        {
            var item = review.Captures.SingleOrDefault(c => c.Id == capture.Id);
            if (item is null || item.Sha256 != capture.Sha256 || item.Verdict is not ("ok" or "intentional-change" or "defect")
                || string.IsNullOrWhiteSpace(item.Note)) throw new InvalidDataException("Missing verdict / reason / matching hash: " + capture.Id);
        }
        string status = review.Captures.Any(c => c.Verdict == "defect") ? "REJECTED" : "REVIEWED";
        string comparison = Path.Combine(root, "comparison/comparison.json");
        using JsonDocument development = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "development.json")));
        if (development.RootElement.GetProperty("functionalStatus").GetString() != "PASS") throw new InvalidDataException("Functional session failed");
        DevFiles.Write(Path.Combine(root, "visual-review.json"), new { status, reviewedAtUtc = DateTimeOffset.UtcNow, review,
            comparisonSha256 = File.Exists(comparison) ? DevFiles.Hash(comparison) : null,
            scope = "Records review decisions for exact captures; hardware behavior and completion of visual inspection are not independently verified" });
        DevFiles.Write(Path.Combine(root, "development.json"), new { status, functionalStatus = "PASS", visualReview = "visual-review.json",
            sessionSha256 = DevFiles.Hash(session), report = "index.html", hardwareValidation = "NOT_PERFORMED" });
        return status;
    }
}
