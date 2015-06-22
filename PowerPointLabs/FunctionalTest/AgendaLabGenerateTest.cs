using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabGenerateTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesVisualDefault.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabTest()
        {
            TestRemoveAgenda();
            TestGenerateTextAgenda();
            TestGenerateBeamAgenda();
            TestGenerateVisualAgenda();
        }

        private static void TestGenerateVisualAgenda()
        {
            MessageBoxUtil.ExpectMessageBoxWillPopUp("Confirm Update",
                "Agenda already exists. By confirm this dialog agenda will be regenerated. Do you want to proceed?",
                PplFeatures.GenerateVisualAgenda, buttonNameToClick: "OK");

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesVisualDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateBeamAgenda()
        {
            PpOperations.SelectSlide(1);
            MessageBoxUtil.ExpectMessageBoxWillPopUp("Confirm Update",
                "Agenda already exists. By confirm this dialog agenda will be regenerated. Do you want to proceed?",
                PplFeatures.GenerateBeamAgenda, buttonNameToClick: "OK");

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesBeamDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateTextAgenda()
        {
            PplFeatures.GenerateTextAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesTextDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestRemoveAgenda()
        {
            PplFeatures.RemoveAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }
    }
}
