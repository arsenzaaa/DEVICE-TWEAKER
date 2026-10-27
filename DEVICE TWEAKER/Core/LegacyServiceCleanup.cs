namespace DeviceTweakerCS;

// Preserve the published ownership rules; never stop a service merely because it exists.
internal static class LegacyServiceCleanup
{
    internal static void Run(bool created, bool started, bool initialized,
        Action close, Action stopOwnedService, Action cleanupOwnedArtifacts, Action release)
    {
        try
        {
            close();
            if (created || started) stopOwnedService();
            if (initialized && (created || started)) cleanupOwnedArtifacts();
        }
        finally { release(); }
    }
}
