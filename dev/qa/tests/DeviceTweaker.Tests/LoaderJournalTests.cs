using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public sealed class LoaderJournalTests : IDisposable
{
    private const string Boot = "11111111-1111-1111-1111-111111111111";
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "DeviceTweakerJournal-" + Guid.NewGuid().ToString("N"));
    private string Journal => Path.Combine(_directory, "loader.pending");
    public LoaderJournalTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, recursive: true);

    [Fact]
    public void PendingFileSurvivesFailureAndPreventsSecondLoadUntilExplicitCompletion()
    {
        DeviceTweakerLoaderJournal.BeginLoaderAt(Journal, Boot);
        Assert.Equal(Boot + "|pending", File.ReadAllText(Journal));
        Assert.Throws<InvalidOperationException>(() => DeviceTweakerLoaderJournal.AssertLoaderIdleAt(Journal, Boot));
        Assert.Throws<InvalidOperationException>(() => DeviceTweakerLoaderJournal.BeginLoaderAt(Journal, Boot));
        DeviceTweakerLoaderJournal.CompleteLoaderAt(Journal);
        DeviceTweakerLoaderJournal.AssertLoaderIdleAt(Journal, Boot);
        Assert.False(File.Exists(Journal));
    }

    [Theory]
    [InlineData("")]
    [InlineData("broken")]
    [InlineData("11111111-1111-1111-1111-111111111111|done")]
    public void EmptyOrMalformedJournalRequiresRecovery(string record)
    {
        File.WriteAllText(Journal, record);
        Assert.Throws<InvalidOperationException>(() => DeviceTweakerLoaderJournal.AssertLoaderIdleAt(Journal, Boot));
        Assert.Equal(record, File.ReadAllText(Journal));
    }

    [Fact]
    public void ValidPendingFileFromPreviousBootIsRemoved()
    {
        DeviceTweakerLoaderJournal.BeginLoaderAt(Journal, Boot);
        DeviceTweakerLoaderJournal.AssertLoaderIdleAt(Journal, "22222222-2222-2222-2222-222222222222");
        Assert.False(File.Exists(Journal));
    }

    [Fact]
    public void PublicSharedEntryPointsRetainQaWriteBarrier()
    {
        Assert.Throws<QaWriteBlockedException>(() => DeviceTweakerLoaderJournal.AssertLoaderIdle());
        Assert.Throws<QaWriteBlockedException>(() => DeviceTweakerLoaderJournal.BeginLoader());
        Assert.Throws<QaWriteBlockedException>(() => DeviceTweakerLoaderJournal.CompleteLoader());
    }
}
