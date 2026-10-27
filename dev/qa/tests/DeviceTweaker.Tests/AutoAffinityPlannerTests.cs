using DeviceTweakerCS;
using Xunit;
namespace DeviceTweaker.Tests;

public sealed class AutoAffinityPlannerTests
{
    private static List<AutoCpuUnit> Cpu(int count, bool smt = true, bool hybrid = false, bool dual = false)
        => Enumerable.Range(0, count).Select(i => new AutoCpuUnit
        { Id = i, PrimaryLp = smt ? i * 2 : i, Lps = smt ? [i * 2, i * 2 + 1] : [i],
            IsECore = hybrid && i >= count / 2, Ccd = dual ? i / Math.Max(1, count / 2) : 0,
            Ccx = dual ? i / Math.Max(1, count / 2) : 0, Rating = 200 - i, Rank = i }).ToList();
    private static List<AutoAffinityPlanSlot> Slots(params AutoAffinityRole[] roles)
        => roles.Select((r, i) => new AutoAffinityPlanSlot { Kind = r == AutoAffinityRole.Gpu ? DeviceKind.GPU : r == AutoAffinityRole.Nic ? DeviceKind.NET_NDIS : DeviceKind.USB,
            InstanceId = $"device-{i}", Role = r, Need = r == AutoAffinityRole.Gpu ? 2 : 1, OriginalIndex = i }).ToList();
    private static AutoAllocationResult Plan(List<AutoAffinityPlanSlot> slots, List<AutoCpuUnit> units, bool dual = false)
    {
        var preferred = units.Where(u => !u.IsECore && (!dual || u.Ccd == 0)).ToList();
        var background = units.Where(u => !u.HasCore0).Reverse().ToList();
        return AutoAffinityPlanner.Allocate(slots, preferred, background, units,
            dual ? preferred.SelectMany(u => u.Lps).ToList() : [], units.Sum(u => u.Lps.Count), dual, _ => "", _ => { });
    }
    [Fact] public void GpuKeepsHighestAdjacentPhysicalPairDespiteCppcOrder()
    {
        var units = Cpu(8); // ordering supplied by the caller
        units.Reverse(); var slots = Slots(AutoAffinityRole.InputMouse, AutoAffinityRole.Gpu);
        // Input uses the front of the provided order. Use the actual ranked order
        // for the baseline; GPU must choose the tail based on physical topology.
        units = units.OrderBy(u => u.PrimaryLp == 6 ? 0 : u.PrimaryLp + 1).ToList();
        Plan(slots, units); Assert.Equal(new[] { 12, 14 }, slots[1].Lps); Assert.DoesNotContain(6, slots[1].Lps);
    }
    [Fact] public void KeyboardCanShareInputSiblingWhenDedicatedCoresRunOut()
    {
        var slots = Slots(AutoAffinityRole.InputMouse, AutoAffinityRole.Keyboard); Plan(slots, Cpu(1));
        Assert.Equal(new[] { 0 }, slots[0].Lps); Assert.Equal(new[] { 1 }, slots[1].Lps);
    }
    [Fact] public void GpuWithoutTwoPhysicalCoresRemainsUnassigned()
    {
        var slots = Slots(AutoAffinityRole.Gpu); Plan(slots, Cpu(1)); Assert.Empty(slots[0].Lps);
    }
    [Theory]
    [InlineData(2, false, false, false)] [InlineData(4, true, false, false)]
    [InlineData(8, true, false, false)] [InlineData(12, false, true, false)]
    [InlineData(16, true, false, true)] [InlineData(24, false, true, true)]
    public void AllocationRespectsPhysicalIsolationAndValidLogicalProcessors(int cores, bool smt, bool hybrid, bool dual)
    {
        var units = Cpu(cores, smt, hybrid, dual); var slots = Slots(AutoAffinityRole.InputMouse, AutoAffinityRole.InputController,
            AutoAffinityRole.Gpu, AutoAffinityRole.Nic, AutoAffinityRole.Keyboard, AutoAffinityRole.Audio, AutoAffinityRole.OtherUsb);
        var result = Plan(slots, units, dual); var byLp = units.SelectMany(u => u.Lps.Select(lp => (lp, u))).ToDictionary(p => p.lp, p => p.u);
        Assert.All(slots.SelectMany(s => s.Lps), lp => Assert.True(byLp.ContainsKey(lp)));
        var gpu = slots.Single(s => s.Role == AutoAffinityRole.Gpu);
        if (gpu.Lps.Count > 0) { Assert.Equal(2, gpu.Lps.Count); Assert.Equal(2, gpu.Lps.Select(lp => byLp[lp].Id).Distinct().Count()); Assert.All(gpu.Lps, lp => Assert.False(byLp[lp].IsECore)); }
        var gpuCores = gpu.Lps.Select(lp => byLp[lp].Id).ToHashSet();
        Assert.All(slots.Where(s => s != gpu).SelectMany(s => s.Lps), lp => Assert.DoesNotContain(byLp[lp].Id, gpuCores));
        var inputCores = slots.Where(s => s.Role is AutoAffinityRole.InputMouse or AutoAffinityRole.InputController).SelectMany(s => s.Lps).Select(lp => byLp[lp].Id).ToHashSet();
        Assert.All(slots.Where(s => s.Role is AutoAffinityRole.Nic or AutoAffinityRole.Audio).SelectMany(s => s.Lps), lp => Assert.DoesNotContain(byLp[lp].Id, inputCores));
        Assert.All(slots.SelectMany(s => s.Lps), lp => Assert.Contains(lp, result.ConsumedCores));
    }
}
