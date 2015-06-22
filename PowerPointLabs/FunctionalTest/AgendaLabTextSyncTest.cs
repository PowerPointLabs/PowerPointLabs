using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabTextSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesTextBeforeSync.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabTextSyncTest()
        {
            PplFeatures.SynchronizeAgenda();
            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesTextAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
