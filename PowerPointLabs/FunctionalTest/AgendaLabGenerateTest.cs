using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabGenerateTest : BaseFunctionalTest
    {
        private const string AgendaExistsTitle = "Confirm Update";
        private const string AgendaExistsContent =
            "Agenda already exists. By confirm this dialog agenda will be regenerated. Do you want to proceed?";

        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesVisualDefault.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabGenerateTest()
        {
            TestRemoveAgenda();
            NoAgendaRemoveUnsuccessful();

            TestGenerateTextAgenda();
            TestGenerateBeamAgenda();
            TestGenerateVisualAgenda();
            
            EmptySectionUnsuccessful(false);
            LongSectionNameUnsuccessful(false);
            OneSectionUnsuccessful(false);
            NoSectionUnsuccessful(false);
        }

        [TestMethod]
        public void FT_AgendaLabInvalidSectionSyncTest()
        {
            EmptySectionUnsuccessful(true);
            LongSectionNameUnsuccessful(true);
            OneSectionUnsuccessful(true);
            NoSectionUnsuccessful(true);
        }

        private static void TestGenerateVisualAgenda()
        {
            MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
                PplFeatures.GenerateVisualAgenda, buttonNameToClick: "OK");

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesVisualDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateBeamAgenda()
        {
            PpOperations.SelectSlide(1);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
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


        private void NoAgendaRemoveUnsuccessful()
        {
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "There's no generated agenda.",
                PplFeatures.RemoveAgenda);
        }

        public void EmptySectionUnsuccessful(bool isTestingSynchronize)
        {
            var slide = PpOperations.SelectSlide(27);
            slide.Delete();

            var title = "Unable to execute action";
            var message = "Presentation contains empty section(s). Please fill them up or remove them.";

            if (isTestingSynchronize)
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.SynchronizeAgenda);
            }
            else
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, () =>
                    MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
                        PplFeatures.GenerateBeamAgenda, buttonNameToClick: "OK"));
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateVisualAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
            }

            PpOperations.DeleteSection(6, true);
        }

        public void LongSectionNameUnsuccessful(bool isTestingSynchronize)
        {
            string longName = new string('x', 200);
            PpOperations.RenameSection(2, longName);

            var title = "Unable to execute action";
            var message = "One of the section names exceeds the maximum size allowed by Agenda Lab. Please rename the section accordingly.";

            if (isTestingSynchronize)
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.SynchronizeAgenda);
            }
            else
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateVisualAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
            }

            PpOperations.RenameSection(2, "One Slide");
        }

        public void OneSectionUnsuccessful(bool isTestingSynchronize)
        {
            PpOperations.DeleteSection(5, false);
            PpOperations.DeleteSection(4, false);
            PpOperations.DeleteSection(3, false);
            PpOperations.DeleteSection(2, false);

            var title = "Unable to execute action";
            var message = "Agenda Lab requires slides to be divided into two or more sections.";

            if (isTestingSynchronize)
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.SynchronizeAgenda);
            }
            else
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateVisualAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
            }
        }

        public void NoSectionUnsuccessful(bool isTestingSynchronize)
        {
            PpOperations.DeleteSection(1, false);

            var title = "Unable to execute action";
            var message = "Please group slides into sections before generating agenda.";

            if (isTestingSynchronize)
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.SynchronizeAgenda);
            }
            else
            {
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateVisualAgenda);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(title, message, PplFeatures.GenerateBeamAgenda);
            }
        }
    }
}
