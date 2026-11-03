namespace DeviceTweakerCS;

// The production restore entry point and its file-backed regressions use this
// order: validate the chosen document, create rollback, restore the prepared data.
internal sealed class BackupRestoreWorkflow(
    Func<string, DeviceTweakerBackup> read,
    Func<string, bool> createRollback,
    Func<DeviceTweakerBackup, OperationReport> restore)
{
    internal OperationReport Restore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("A backup path is required.", nameof(path));
        DeviceTweakerBackup prepared = read(path);
        if (!createRollback(path))
            throw new InvalidOperationException("A rollback backup could not be created; restore was cancelled.");
        return restore(prepared);
    }
}

internal sealed record BackupFileEntry(string Path, DateTime LastWriteTimeUtc);

internal static class BackupFileRetention
{
    internal const string OriginalFileName = "DeviceTweakerBackup_ORIGINAL.json";

    internal static IReadOnlyList<string> SelectForRemoval(
        IEnumerable<BackupFileEntry> candidates, int keepLatest, IEnumerable<string>? protectedPaths = null)
    {
        var files = candidates
            .Where(file => !string.Equals(System.IO.Path.GetFileName(file.Path),
                OriginalFileName, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .ThenBy(file => file.Path, StringComparer.OrdinalIgnoreCase).ToArray();
        var protectedSet = new HashSet<string>((protectedPaths ?? [])
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(System.IO.Path.GetFullPath), StringComparer.OrdinalIgnoreCase);
        var kept = new HashSet<string>(files.Where(file =>
            protectedSet.Contains(System.IO.Path.GetFullPath(file.Path)))
            .Select(file => file.Path), StringComparer.OrdinalIgnoreCase);
        int limit = Math.Max(Math.Max(1, keepLatest), kept.Count);
        var removed = new List<string>();
        foreach (var file in files)
        {
            if (kept.Contains(file.Path)) continue;
            if (kept.Count < limit) kept.Add(file.Path);
            else removed.Add(file.Path);
        }
        return removed;
    }
}
