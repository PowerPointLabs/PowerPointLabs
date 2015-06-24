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
            TextSyncSuccessful();
            NoContentShapeUnsuccessful();
            NoRefSlideUnsuccessful();
            NoAgendaUnsuccessful();
        }

        public void TextSyncSuccessful() {
            PplFeatures.SynchronizeAgenda();
            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesTextAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }


        public void NoContentShapeUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            var contentShape = PpOperations.SelectShapesByPrefix("PptLabsAgenda_&^@ContentShape")[1];
            contentShape.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The reference slide is invalid. Please remove and regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoRefSlideUnsuccessful()
        {
            var refSlide = PpOperations.SelectSlide(1);
            refSlide.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The reference slide is missing. Please remove and regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoAgendaUnsuccessful()
        {
            PplFeatures.RemoveAgenda();
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "There's no generated agenda.",
                PplFeatures.SynchronizeAgenda);
        }
    }
}
