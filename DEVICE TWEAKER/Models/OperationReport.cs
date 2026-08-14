namespace DeviceTweakerCS;

internal sealed class OperationReport
{
    private readonly List<string> _errors = [];

    public IReadOnlyList<string> Errors => _errors;
    public bool Succeeded => _errors.Count == 0;

    public void AddError(string context, string? message)
    {
        string detail = string.IsNullOrWhiteSpace(message) ? "unknown error" : message.Trim();
        string entry = string.IsNullOrWhiteSpace(context) ? detail : $"{context}: {detail}";
        if (!_errors.Contains(entry, StringComparer.OrdinalIgnoreCase))
        {
            _errors.Add(entry);
        }
    }
}
