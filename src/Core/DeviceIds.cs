namespace DeviceTweakerCS;

internal static class DeviceIds
{
    // Instance IDs compare case-insensitively; some sources escape the backslashes.
    internal static string Normalize(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return string.Empty;
        }

        return id.Replace("\\\\", "\\").Trim().ToUpperInvariant();
    }
}
