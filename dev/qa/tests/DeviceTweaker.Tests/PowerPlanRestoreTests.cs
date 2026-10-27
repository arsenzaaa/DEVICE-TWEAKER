using DeviceTweakerCS;
using Xunit;
namespace DeviceTweaker.Tests;

public sealed class PowerPlanRestoreTests
{
    private sealed class Backend : IPowerPlanBackend
    {
        internal Guid Active = Guid.NewGuid();
        internal readonly List<(Guid Scheme, bool Ac, uint Value)> Writes = [];
        internal readonly List<Guid> Activations = [];
        internal readonly Dictionary<(Guid, bool), uint> Values = [];
        internal string? Fault;
        public Guid GetActiveScheme() => Fault == "active" ? Guid.Empty : Active;
        public uint ReadIndex(Guid s, bool ac) => Fault == "readback" ? 99u : Values[(s, ac)];
        public void WriteIndex(Guid s, bool ac, uint index)
        { Writes.Add((s, ac, index)); if (Fault == "write") throw new IOException("write failed"); Values[(s, ac)] = index; }
        public void Activate(Guid s) { Activations.Add(s); if (Fault == "activate") throw new IOException("activate failed"); }
    }
    [Theory] [InlineData(true, true)] [InlineData(true, false)] [InlineData(false, true)] [InlineData(false, false)]
    public void RestoreTouchesOnlySavedSchemeAndSelectedPowerSource(bool active, bool ac)
    {
        Backend b = new(); Guid saved = active ? b.Active : Guid.NewGuid(); b.Values[(saved, !ac)] = 1;
        PowerPlanRestoreService.Restore(b, saved, ac, 0);
        Assert.Equal((saved, ac, 0u), Assert.Single(b.Writes)); Assert.Equal(1u, b.Values[(saved, !ac)]);
        Assert.Equal(active ? new[] { saved } : [], b.Activations);
    }
    [Theory] [InlineData("active", 0)] [InlineData("write", 1)] [InlineData("activate", 1)] [InlineData("readback", 1)]
    public void ApiFailureIsNotReportedAsSuccess(string fault, int writes)
    {
        Backend b = new() { Fault = fault };
        Assert.ThrowsAny<Exception>(() => PowerPlanRestoreService.Restore(b, b.Active, true, 0));
        Assert.Equal(writes, b.Writes.Count);
    }
}
