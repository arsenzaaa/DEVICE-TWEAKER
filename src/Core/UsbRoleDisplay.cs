namespace DeviceTweakerCS;

internal static class UsbRoleDisplay
{
    private static readonly string[] RoleOrder = ["Mouse", "Keyboard", "Gamepad", "Microphone", "Audio", "Webcam"];

    internal static int Rank(string role)
    {
        for (int index = 0; index < RoleOrder.Length; index++)
        {
            if (RawPollingService.IsUsbRoleText(role, RoleOrder[index])) return index;
        }

        return RoleOrder.Length;
    }

    internal static IEnumerable<string> Order(IEnumerable<string> roles) => roles
        .Select(role => role.Trim())
        .Where(role => role.Length > 0)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(Rank)
        .ThenBy(role => role, StringComparer.OrdinalIgnoreCase);

    internal static string Format(string roles) => string.Join(", ", Order(
        roles.Split(',', StringSplitOptions.RemoveEmptyEntries)));
}
