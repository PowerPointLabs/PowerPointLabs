using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabBeamSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesBeamBeforeSync.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabBeamSyncTest()
        {
            // TODO: Is there really no way to programmatically select multiple slides at once?
            // TODO: Ideally, I want to select the slides 5,6,7,8 together and apply Synchronise Agenda on them

            PpOperations.SelectSlide(5);
            PplFeatures.SynchronizeAgenda();
            PpOperations.SelectSlide(6);
            PplFeatures.SynchronizeAgenda();
            PpOperations.SelectSlide(8);
            PplFeatures.SynchronizeAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesBeamAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}