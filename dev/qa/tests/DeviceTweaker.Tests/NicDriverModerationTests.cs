namespace DeviceTweaker.Tests;

using DeviceTweakerCS;
using Xunit;

public class NicDriverModerationTests
{
    private static readonly Dictionary<string, string> OnOff = new() { ["0"] = "Disabled", ["1"] = "Enabled" };
    private static readonly Dictionary<string, string> IntelItr = new()
    {
        ["65535"] = "Adaptive", ["3600"] = "Extreme", ["950"] = "Medium", ["0"] = "Off",
    };

    [Fact]
    public void PlanDisable_TurnsOffEveryKeywordThatOffersOff()
    {
        List<NicDriverModeration.Change> changes = NicDriverModeration.PlanDisable(
        [
            new("*InterruptModeration", OnOff, "1", "1"),
            new("ITR", IntelItr, "65535", "65535"),
        ]);

        Assert.Equal(
            [new NicDriverModeration.Change("*InterruptModeration", "1", "0"), new NicDriverModeration.Change("ITR", "65535", "0")],
            changes);
    }

    [Fact]
    public void PlanDisable_SkipsValuesAlreadyOffAndDriversWithoutOff()
    {
        Dictionary<string, string> noOff = new() { ["1"] = "Low", ["2"] = "High" };

        Assert.Empty(NicDriverModeration.PlanDisable(
        [
            new("*InterruptModeration", OnOff, "1", "0"),
            new("ITR", noOff, "1", "2"),
        ]));
    }

    [Fact]
    public void PlanDisable_WritesValueMissingFromClassKey()
    {
        List<NicDriverModeration.Change> changes = NicDriverModeration.PlanDisable([new("*InterruptModeration", OnOff, "1", null)]);

        Assert.Equal([new NicDriverModeration.Change("*InterruptModeration", null, "0")], changes);
    }

    [Fact]
    public void PlanRestoreDefault_RestoresOnlyValuesStillOff()
    {
        List<NicDriverModeration.Change> changes = NicDriverModeration.PlanRestoreDefault(
        [
            new("*InterruptModeration", OnOff, "1", "0"),
            new("ITR", IntelItr, "65535", "950"),
        ]);

        Assert.Equal([new NicDriverModeration.Change("*InterruptModeration", "0", "1")], changes);
    }

    [Fact]
    public void PlanRestoreDefault_KeepsOffWhenDriverDefaultIsOff()
    {
        Assert.Empty(NicDriverModeration.PlanRestoreDefault([new("*InterruptModeration", OnOff, "0", "0")]));
    }

    [Fact]
    public void Describe_UsesDriverOptionText()
    {
        string? text = NicDriverModeration.Describe(
        [
            new("*InterruptModeration", OnOff, "1", "1"),
            new("ITR", IntelItr, "65535", null),
        ]);

        Assert.Equal("Enabled, ITR Adaptive", text);
    }

    [Fact]
    public void Describe_ReturnsNullWithoutDriverKeywords()
    {
        Assert.Null(NicDriverModeration.Describe([]));
    }

    [Theory]
    [InlineData("*InterruptModeration")]
    [InlineData("ITR")]
    public void BackupAcceptsModerationValuesInAdapterClassKey(string name)
    {
        RegistryValueBackup value = new()
        {
            Hive = "HKLM",
            Path = @"SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\0001",
            Name = name,
            Exists = true,
            Kind = "String",
            Data = "0",
        };

        Assert.True(BackupFormat.IsManagedBackupRegistryValue(value));
    }
}
