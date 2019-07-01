using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabTextSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaLab\\AgendaSlidesTextBeforeSync.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabTextSyncTest()
        {
            HideUnvisitedSyncSuccessful();
            TextSyncSuccessful();
            NoContentShapeUnsuccessful();
            NoRefSlideUnsuccessful();
            NoAgendaUnsuccessful();
        }

        public void HideUnvisitedSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();

            // Duplicate template slide and delete original template slide. It should use the duplicate as the new template slide.
            Microsoft.Office.Interop.PowerPoint.Slide firstSlide = PpOperations.SelectSlide(1);

            PpOperations.SelectShape("PptLabsAgenda_&^@ContentShape_&^@2015061916283877850").TextFrame2.TextRange.Paragraphs[3].Text = " ";

            PplFeatures.SynchronizeAgenda();

            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesTextAfterSyncHideUnvisited.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);

            PpOperations.SelectShape("PptLabsAgenda_&^@ContentShape_&^@2015061916283877850").TextFrame2.TextRange.Paragraphs[3].Text = "Readd bullet format";

            PplFeatures.SynchronizeAgenda();

            actualSlides = PpOperations.FetchCurrentPresentationData();
            expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesTextAfterSyncUnhideUnvisited.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);

        }


        public void TextSyncSuccessful()
        {
            PplFeatures.SynchronizeAgenda();

            // Duplicate template slide and delete original template slide. It should use the duplicate as the new template slide.
            Microsoft.Office.Interop.PowerPoint.Slide firstSlide = PpOperations.SelectSlide(1);
            PpOperations.SelectShape("PPTTemplateMarker").SafeDelete();
            firstSlide.Duplicate();
            firstSlide.Delete();

            PplFeatures.SynchronizeAgenda();

            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesTextAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }


        public void NoContentShapeUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            Microsoft.Office.Interop.PowerPoint.Shape contentShape = PpOperations.SelectShapesByPrefix("PptLabsAgenda_&^@ContentShape")[1];
            contentShape.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoRefSlideUnsuccessful()
        {
            Microsoft.Office.Interop.PowerPoint.Slide refSlide = PpOperations.SelectSlide(1);
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
