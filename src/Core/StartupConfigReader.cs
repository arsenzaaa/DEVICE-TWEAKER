using System.Text.Json;

namespace DeviceTweakerCS;

internal enum StartupConfigReadState { Missing, Loaded, Error }

internal sealed record StartupConfigReadResult<T>(
    StartupConfigReadState State, T? Value, string? Source, Exception? Error) where T : class
{
    internal T RequireValueOrCreate(Func<T> create)
    {
        if (State == StartupConfigReadState.Error)
            throw new InvalidDataException($"The saved startup profile could not be read: {Source}. No settings were saved.", Error);
        return Value ?? create();
    }
}

// Only a missing file permits trying an older layout. Access, I/O and parse
// failures must never turn an existing profile into a new, empty one.
internal static class StartupConfigReader
{
    internal static StartupConfigReadResult<T> Read<T>(
        string profilePath, IEnumerable<string> legacyPaths,
        Func<string, T> readCurrent, Func<string, T> readLegacy) where T : class
    {
        var paths = new[] { profilePath }.Concat(legacyPaths).ToArray();
        for (int index = 0; index < paths.Length; index++)
        {
            string path = paths[index];
            try
            {
                T value = index == 0 ? readCurrent(path) : readLegacy(path);
                return new(StartupConfigReadState.Loaded, value, path, null);
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException
                or InvalidDataException or JsonException or FormatException or OverflowException)
            {
                return new(StartupConfigReadState.Error, null, path, error);
            }
        }
        return new(StartupConfigReadState.Missing, null, null, null);
    }
}
