using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.AudioMisc;

namespace Test.FunctionalTest
{
    [TestClass]
    public class NarrationsLabTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "NarrationsLab\\NarrationsLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_NarrationsLegacyTest()
        {
            // ensure that the audio type detection is
            // backwards-compatible with older versions
            Shape shape = PpOperations.SelectShape("Record")[1];
            Assert.AreEqual(Audio.AudioType.Record, Audio.GetShapeAudioType(shape));
            shape = PpOperations.SelectShape("Auto")[1];
            Assert.AreEqual(Audio.AudioType.Auto, Audio.GetShapeAudioType(shape));
            shape = PpOperations.SelectShape("Unrecognized")[1];
            Assert.AreEqual(Audio.AudioType.Unrecognized, Audio.GetShapeAudioType(shape));
        }
    }
}
