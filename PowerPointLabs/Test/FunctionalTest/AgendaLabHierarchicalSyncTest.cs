using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabHierarchicalSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesTextHierarchicalBeforeSync.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabHierarchicalSyncTest()
        {
            TextSyncSuccessful();
        }

        public void TextSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();

            PplFeatures.SynchronizeAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesTextHierarchicalAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
