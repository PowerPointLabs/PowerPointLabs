using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabVisualSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesVisualBeforeSync.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabVisualSyncTest()
        {
            VisualSyncSuccessful();
        }

        public void VisualSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();
            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesVisualAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
