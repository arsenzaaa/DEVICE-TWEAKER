namespace DeviceTweakerCS;

internal static class ImodDriverPaths
{
    internal static string SystemPath()
    {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrWhiteSpace(windows))
        {
            windows = Path.GetPathRoot(Environment.SystemDirectory) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(windows))
            {
                windows = AppContext.BaseDirectory;
                if (string.IsNullOrWhiteSpace(windows))
                {
                    windows = Environment.CurrentDirectory;
                }
            }
        }

        return Path.Combine(windows, "DTIMOD.sys");
    }

    internal static bool IsSystemPath(string? path)
    {
        string systemPath = SystemPath();
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(systemPath))
        {
            return false;
        }

        try
        {
            return string.Equals(Path.GetFullPath(path), Path.GetFullPath(systemPath), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(path, systemPath, StringComparison.OrdinalIgnoreCase);
        }
    }
}
