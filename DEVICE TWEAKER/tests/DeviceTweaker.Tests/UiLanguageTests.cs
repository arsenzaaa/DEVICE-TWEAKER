using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class UiLanguageTests
{
    [Fact]
    public void RussianTranslationContract_PreservesExpectedTerms()
    {
        Assert.True(UiLanguage.ValidateRussianTranslationContract(out string error), error);
    }
}
