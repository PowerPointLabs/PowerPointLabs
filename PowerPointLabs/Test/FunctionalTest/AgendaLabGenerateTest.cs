using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabGenerateTest : BaseFunctionalTest
    {
        private const string AgendaExistsTitle = "Confirm Update";
        private const string AgendaExistsContent =
            "Agenda already exists. The previous agenda will be removed and regenerated. Do you want to proceed?";

        protected override string GetTestingSlideName()
        {
            return "AgendaLab\\AgendaSlidesVisualDefault.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
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
        [TestCategory("FT")]
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
            ThreadUtil.WaitFor(10000); // TODO: Remove delay when it becomes more stable
            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesVisualDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateBeamAgenda()
        {
            PpOperations.SelectSlide(1);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(AgendaExistsTitle, AgendaExistsContent,
                PplFeatures.GenerateBeamAgenda, buttonNameToClick: "OK");

            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesBeamDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestGenerateTextAgenda()
        {
            PplFeatures.GenerateTextAgenda();

            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesTextDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        private static void TestRemoveAgenda()
        {
            PplFeatures.RemoveAgenda();

            System.Collections.Generic.List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            System.Collections.Generic.List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesDefault.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }


        private void NoAgendaRemoveUnsuccessful()
        {
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "There is no generated agenda.",
                PplFeatures.RemoveAgenda);
        }

        public void EmptySectionUnsuccessful(bool isTestingSynchronize)
        {
            Microsoft.Office.Interop.PowerPoint.Slide slide = PpOperations.SelectSlide(27);
            slide.Delete();

            string title = "Unable to execute action";
            string message = "Presentation contains empty section(s). Please fill them up or remove them.";

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

            string title = "Unable to execute action";
            string message = "One of the section names exceeds the maximum size allowed by Agenda Lab. Please rename the section accordingly.";

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

            string title = "Unable to execute action";
            string message = "Please divide the slides into two or more sections.";

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

            string title = "Unable to execute action";
            string message = "Please group the slides into sections before generating agenda.";

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
