
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ELearningLab.AudioGenerator;

namespace Test.UnitTest.ELearningLab.Model
{
    [TestClass]
    public class AzureVoiceTest
    {
        private AzureVoice azureVoice;

        [TestInitialize]
        public void Init()
        {
            azureVoice = new AzureVoice(Gender.Female, Locale.enUS, AzureVoiceType.JessaRUS);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestLocalMapping()
        {
            string locale_actual = azureVoice.Locale;
            string locale_expected = "en-US";
            Assert.AreEqual(locale_actual, locale_expected);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestClone()
        {
            AzureVoice voice_expected = new AzureVoice(Gender.Female, Locale.enUS, AzureVoiceType.JessaRUS);
            AzureVoice voice_actual = azureVoice.Clone() as AzureVoice;
            Assert.AreEqual(voice_expected.Rank, voice_actual.Rank);
            Assert.AreEqual(voice_expected.voiceType, voice_actual.voiceType);
            Assert.AreEqual(voice_expected.locale, voice_actual.locale);
            Assert.AreEqual(voice_expected.voiceName, voice_actual.voiceName);
            Assert.AreEqual(voice_expected.VoiceName, voice_actual.VoiceName);
        }

    }
}
