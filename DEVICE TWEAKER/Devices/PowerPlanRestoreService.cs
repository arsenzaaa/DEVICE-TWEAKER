namespace DeviceTweakerCS;

internal interface IPowerPlanBackend
{
    Guid GetActiveScheme();
    uint ReadIndex(Guid scheme, bool ac);
    void WriteIndex(Guid scheme, bool ac, uint index);
    void Activate(Guid scheme);
}

internal static class PowerPlanRestoreService
{
    internal static void Restore(IPowerPlanBackend backend, Guid scheme, bool ac, uint index)
    {
        Guid active = backend.GetActiveScheme();
        if (scheme == Guid.Empty || active == Guid.Empty)
            throw new InvalidOperationException("Could not read the Windows power scheme.");
        backend.WriteIndex(scheme, ac, index);
        if (active == scheme) backend.Activate(scheme);
        uint actual = backend.ReadIndex(scheme, ac);
        if (actual != index)
            throw new InvalidOperationException($"Power setting verification mismatch: expected={index} actual={actual}.");
    }
}
