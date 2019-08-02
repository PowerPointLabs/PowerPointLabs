using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabVisualSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaLab\\AgendaSlidesVisualBeforeSync.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabVisualSyncTest()
        {
            VisualSyncSuccessful();
        }

        public void VisualSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();
            ThreadUtil.WaitFor(10000);
            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesVisualAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
