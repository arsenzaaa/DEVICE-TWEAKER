namespace DeviceTweakerCS;

internal enum OperationIssueSeverity
{
    Success,
    Warning,
    Error,
}

internal sealed record OperationIssue(
    OperationIssueSeverity Severity,
    string Component,
    string? UserMessage,
    string? TechnicalDetails)
{
    public bool IsSkipped { get; init; }
    public string DisplayText => string.IsNullOrWhiteSpace(UserMessage)
        ? Component
        : $"{Component}: {UserMessage}";
}

internal sealed class OperationReport
{
#if DEVELOPER_QA
    internal bool IsDeveloperPreview { get; init; }
#else
    internal bool IsDeveloperPreview => false;
#endif
    private readonly List<OperationIssue> _issues = [];

    public IReadOnlyList<OperationIssue> Issues => _issues;
    public IReadOnlyList<string> Errors => _issues
        .Where(issue => issue.Severity == OperationIssueSeverity.Error)
        .Select(issue => issue.DisplayText)
        .ToList();
    public IReadOnlyList<OperationIssue> Warnings => _issues
        .Where(issue => issue.Severity == OperationIssueSeverity.Warning)
        .ToList();
    public bool Succeeded => !_issues.Any(issue => issue.Severity == OperationIssueSeverity.Error);
    public bool NoChangesMade { get; private set; }
    public string? BackupPath { get; private set; }

    public void MarkNoChangesMade()
    {
        NoChangesMade = true;
    }

    public void SetBackupPath(string? path)
    {
        BackupPath = string.IsNullOrWhiteSpace(path) ? null : path.Trim();
    }

    public void AddSuccess(string component, string? message = null)
    {
        AddIssue(OperationIssueSeverity.Success, component, message, null);
    }

    public void AddWarning(string component, string? userMessage, string? technicalDetails = null)
    {
        AddIssue(OperationIssueSeverity.Warning, component, userMessage, technicalDetails);
    }

    public void AddSkipped(string component, string? userMessage, string? technicalDetails = null)
    {
        AddIssue(OperationIssueSeverity.Warning, component, userMessage, technicalDetails, isSkipped: true);
    }

    public void AddError(string context, string? message)
    {
        // Most existing callers pass exception text. Preserve it for DETAILS,
        // but do not force raw diagnostics into the compact summary.
        AddIssue(OperationIssueSeverity.Error, context, null, message);
    }

    public void AddError(string context, string? userMessage, string? technicalDetails)
    {
        AddIssue(OperationIssueSeverity.Error, context, userMessage, technicalDetails);
    }

    private void AddIssue(
        OperationIssueSeverity severity,
        string component,
        string? userMessage,
        string? technicalDetails,
        bool isSkipped = false)
    {
        string cleanComponent = string.IsNullOrWhiteSpace(component) ? "Operation" : component.Trim();
        string? cleanUserMessage = string.IsNullOrWhiteSpace(userMessage) ? null : userMessage.Trim();
        string? cleanTechnicalDetails = string.IsNullOrWhiteSpace(technicalDetails) ? null : technicalDetails.Trim();

        bool duplicate = _issues.Any(issue =>
            issue.Severity == severity
            && string.Equals(issue.Component, cleanComponent, StringComparison.OrdinalIgnoreCase)
            && string.Equals(issue.UserMessage, cleanUserMessage, StringComparison.OrdinalIgnoreCase)
            && string.Equals(issue.TechnicalDetails, cleanTechnicalDetails, StringComparison.OrdinalIgnoreCase)
            && issue.IsSkipped == isSkipped);
        if (!duplicate)
        {
            _issues.Add(new OperationIssue(severity, cleanComponent, cleanUserMessage, cleanTechnicalDetails) { IsSkipped = isSkipped });
        }
    }
}
