using System.Text;

namespace DeviceTweakerCS;

// Writes a complete temporary file next to the target, flushes it and moves it
// over the target, so readers never see a partial file. A read-only QA process
// may write only inside its run directory.
internal static class AtomicFile
{
    internal static void WriteAllText(string path, string content, Encoding encoding)
    {
        Write(path, stream =>
        {
            using StreamWriter writer = new(stream, encoding, bufferSize: 4096, leaveOpen: true);
            writer.Write(content);
            writer.Flush();
        });
    }

    internal static void WriteAllBytes(string path, byte[] content)
        => Write(path, stream => stream.Write(content));

    private static void Write(string path, Action<FileStream> writeContent)
    {
        QaRun.DemandArtifactPath(path);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory))
        {
            throw new InvalidOperationException($"Cannot resolve the destination directory for {path}.");
        }

        Directory.CreateDirectory(directory);
        string temporaryPath = Path.Combine(directory, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (FileStream stream = new(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       bufferSize: 4096,
                       FileOptions.WriteThrough))
            {
                writeContent(stream);
                stream.Flush(flushToDisk: true);
            }

            File.Move(temporaryPath, fullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }
}
