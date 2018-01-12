using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using System.Diagnostics;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabWithSlideNumberGenerateTest : BaseFunctionalTest
    {
        private const string AgendaExistsTitle = "Confirm Update";
        private const string AgendaExistsContent =
            "Agenda already exists. The previous agenda will be removed and regenerated. Do you want to proceed?";

        protected override string GetTestingSlideName()
        {
            return "AgendaLab\\AgendaSlidesDefault.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabWithSlideNumberGenerateTest()
        {
            TestGenerateTextAgendaWithSlideNumbers();
            TestGenerateBeamAgendaWithSlideNumbers();
            TestGenerateVisualAgendaWithSlideNumbers();
        }

        private static void TestGenerateTextAgendaWithSlideNumbers()
        {
            PpOperations.ShowAllSlideNumbers();
            PplFeatures.GenerateTextAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesTextWithSlideNumberDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateBeamAgendaWithSlideNumbers()
        {
            PpOperations.ShowAllSlideNumbers();
            PpOperations.SelectSlide(1);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
                PplFeatures.GenerateBeamAgenda, buttonNameToClick: "OK");

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesBeamWithSlideNumberDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateVisualAgendaWithSlideNumbers()
        {
            PpOperations.ShowAllSlideNumbers();
            MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
                PplFeatures.GenerateVisualAgenda, buttonNameToClick: "OK");

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesVisualWithSlideNumberDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
