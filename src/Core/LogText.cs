namespace DeviceTweakerCS;

internal static class LogText
{
    // Keeps one log record on one line: exception text and paths can contain line breaks.
    internal static string Flatten(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value
            .Replace("\r\n", " | ")
            .Replace("\n", " | ")
            .Replace("\r", " | ")
            .Trim();
    }
}
