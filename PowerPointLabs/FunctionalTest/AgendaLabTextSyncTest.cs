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
        [TestCategory("FT")]
        public void FT_AgendaLabTextSyncTest()
        {
            TextSyncSuccessful();
            NoContentShapeUnsuccessful();
            NoRefSlideUnsuccessful();
            NoAgendaUnsuccessful();
        }

        public void TextSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();

            // Duplicate template slide and delete original template slide. It should use the duplicate as the new template slide.
            var firstSlide = PpOperations.SelectSlide(1);
            PpOperations.SelectShape("PPTTemplateMarker").Delete();
            firstSlide.Duplicate();
            firstSlide.Delete();

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
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoRefSlideUnsuccessful()
        {
            var refSlide = PpOperations.SelectSlide(1);
            refSlide.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "No reference slide could be found. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoAgendaUnsuccessful()
        {
            PplFeatures.RemoveAgenda();
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "There is no generated agenda.",
                PplFeatures.SynchronizeAgenda);
        }
    }
}
