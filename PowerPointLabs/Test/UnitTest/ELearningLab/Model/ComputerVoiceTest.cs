using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ELearningLab.AudioGenerator;

namespace Test.UnitTest.ELearningLab.Model
{
    [TestClass]
    public class ComputerVoiceTest
    {
        private ComputerVoice computerVoice;

        [TestInitialize]
        public void Init()
        {
            computerVoice = new ComputerVoice("computer voice");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestClone()
        {
            ComputerVoice voice_expected = new ComputerVoice("computer voice");
            ComputerVoice voice_actual = computerVoice.Clone() as ComputerVoice;
            Assert.AreEqual(voice_expected.Rank, voice_actual.Rank);
            Assert.AreEqual(voice_expected.Voice, voice_actual.Voice);
        }

    }
}
