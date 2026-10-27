using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class DriverEnvironmentTests
{
    private static DriverServiceState Absent(string name) => new(name, false, "not found");
    private static DriverEnvironmentSnapshot Snapshot(uint? blocklist = 0, bool? hvci = false) => new(
        new(blocklist, blocklist.HasValue ? "registry" : "missing"), new(0, "registry"), hvci,
        new("WinDefend", true, "stopped"), Absent("FACEIT"), Absent("vgk"), Absent("vgc"));

    [Fact] public void MissingValueIsUnknownRatherThanEnabledOrDisabled()
    {
        var snapshot = Snapshot(null);
        Assert.Null(snapshot.Blocklist.Enabled);
        Assert.False(snapshot.PotentialRestriction);
        Assert.Contains("not explicitly configured", DriverEnvironment.BuildConfirmation(snapshot));
        Assert.DoesNotContain("Windows blocked", DriverEnvironment.BuildConfirmation(snapshot));
    }
    [Theory]
    [InlineData("FACEIT", "stopped")]
    [InlineData("FACEIT", "running")]
    [InlineData("vgk", "running")]
    [InlineData("vgc", "stopped")]
    public void PresentAntiCheatWarnsWithoutClaimingAnActualConflict(string name, string state)
    {
        var snapshot = Snapshot();
        var service = new DriverServiceState(name, true, state);
        snapshot = name switch { "FACEIT" => snapshot with { Faceit = service }, "vgk" => snapshot with { VanguardDriver = service }, _ => snapshot with { VanguardService = service } };
        string text = DriverEnvironment.BuildConfirmation(snapshot);
        Assert.True(snapshot.AntiCheatDetected);
        Assert.Contains($"found ({state})", text);
        Assert.Contains("a conflict is possible, not detected", text);
        Assert.DoesNotContain("blocked by", text);
    }
    [Fact] public void UnreadableServiceDoesNotBecomeAbsent()
    {
        var snapshot = Snapshot() with { Faceit = new("FACEIT", null, "unavailable (5)") };
        Assert.Contains("FACEIT service: unavailable", DriverEnvironment.BuildConfirmation(snapshot));
        Assert.False(snapshot.AntiCheatDetected);
    }
    [Fact] public void RuntimeAndConfiguredHvciAreKeptSeparate()
    {
        var snapshot = Snapshot(0, true);
        Assert.True(snapshot.PotentialRestriction);
        Assert.Contains("runtime enabled, registry 0 (configured off)", DriverEnvironment.BuildConfirmation(snapshot));
    }
    [Fact] public void DefenderServicePresenceDoesNotProveBlocking()
    {
        var snapshot = Snapshot() with { Defender = new("WinDefend", true, "running") };
        Assert.False(snapshot.PotentialRestriction);
        Assert.Contains("Microsoft Defender service: found (running)", DriverEnvironment.BuildConfirmation(snapshot));
        Assert.DoesNotContain("blocked", DriverEnvironment.BuildConfirmation(snapshot));
    }

    [Fact] public void ActualPreflightMessageTranslatesItsDynamicServiceAndRegistryStates()
    {
        var snapshot = Snapshot(null, true) with { Faceit = new("FACEIT", true, "running") };
        var translate = typeof(UiLanguage).GetMethod("TranslateRussian", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        string text = (string)translate.Invoke(null, [DriverEnvironment.BuildConfirmation(snapshot)])!;
        Assert.Contains("Проверка системы (только чтение)", text);
        Assert.Contains("Vulnerable Driver Blocklist (реестр): нет явного параметра", text);
        Assert.Contains("сейчас включена, в реестре 0 (задано выкл)", text);
        Assert.Contains("Служба FACEIT: найдена (работает)", text);
        Assert.Contains("Vanguard: vgk не найдена, vgc не найдена", text);
        Assert.Contains("конфликт возможен, но не обнаружен", text);
        Assert.DoesNotContain("found (", text);
    }

    [Fact] public void CompactPromptKeepsTheChoiceShortWithoutInventingAConcern()
    {
        string text = DriverEnvironment.BuildCompactConfirmation(Snapshot());
        Assert.StartsWith("Include USB IMOD in auto-optimization?", text);
        Assert.Contains("SKIP continues auto-optimization without USB IMOD.", text);
        Assert.DoesNotContain("SYSTEM CHECK:", text);
        Assert.DoesNotContain("KDU", text);
        Assert.DoesNotContain("conflict", text);
        Assert.DoesNotContain("prevent", text);
    }

    [Theory]
    [InlineData("FACEIT", "stopped")]
    [InlineData("FACEIT", "running")]
    [InlineData("vgk", "running")]
    [InlineData("vgc", "stopped")]
    public void CompactPromptKeepsAntiCheatConcernVisibleWithoutDumpingServiceStates(string name, string state)
    {
        var snapshot = Snapshot();
        var service = new DriverServiceState(name, true, state);
        snapshot = name switch { "FACEIT" => snapshot with { Faceit = service }, "vgk" => snapshot with { VanguardDriver = service }, _ => snapshot with { VanguardService = service } };
        string text = DriverEnvironment.BuildCompactConfirmation(snapshot);
        Assert.Contains(name == "FACEIT" ? "FACEIT" : "Vanguard", text);
        Assert.Contains("may conflict with DTIMOD.sys", text);
        Assert.DoesNotContain("KDU", text);
        Assert.DoesNotContain($"({state})", text);
        Assert.DoesNotContain("registry", text);
    }

    [Theory]
    [InlineData(1u, false)]
    [InlineData(0u, true)]
    public void CompactPromptKeepsObservedWindowsRestrictionVisible(uint blocklist, bool hvci)
    {
        Assert.Contains("Windows settings may prevent DTIMOD.sys from loading", DriverEnvironment.BuildCompactConfirmation(Snapshot(blocklist, hvci)));
    }

    [Fact] public void CompactPromptDoesNotHideFailedReads()
    {
        var snapshot = Snapshot() with { Faceit = new("FACEIT", null, "unavailable (5)") };
        Assert.Contains("Some system checks are unavailable", DriverEnvironment.BuildCompactConfirmation(snapshot));
        Assert.DoesNotContain("FACEIT detected", DriverEnvironment.BuildCompactConfirmation(snapshot));
    }
}
