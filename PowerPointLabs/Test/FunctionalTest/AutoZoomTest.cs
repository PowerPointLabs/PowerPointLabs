using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoZoomTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ZoomLab\\AutoZoom.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoZoomTest()
        {
            // Do tests in reverse order because added slides change slide numbers lower down.
            TestStepBackBackground();
            TestStepBack();
            TestDrillDownBackground();
            TestDrillDown();
            TestDrillDownUnsuccessful();
            TestStepBackUnsuccessful();
        }

        private void TestDrillDown()
        {
            PplFeatures.SetZoomProperties(true, true);

            PpOperations.SelectSlide(4);
            PpOperations.SelectShape("Drill Down This Shape");
            PplFeatures.DrillDown();

            AssertIsSame(4, 7);
            AssertIsSame(5, 8);
            AssertIsSame(6, 9);
        }

        private void TestDrillDownBackground()
        {
            PplFeatures.SetZoomProperties(false, true);

            PpOperations.SelectSlide(10);
            PpOperations.SelectShape("Drill Down This Shape");
            PplFeatures.DrillDown();

            AssertIsSame(10, 13);
            AssertIsSame(11, 14);
            AssertIsSame(12, 15);
        }

        private void TestStepBack()
        {
            PplFeatures.SetZoomProperties(true, true);

            PpOperations.SelectSlide(17);
            PpOperations.SelectShape("Step Back This Shape");
            PplFeatures.StepBack();

            AssertIsSame(16, 20);
            AssertIsSame(17, 21);
            AssertIsSame(18, 22);
        }

        private void TestStepBackBackground()
        {
            PplFeatures.SetZoomProperties(false, true);

            PpOperations.SelectSlide(24);
            PpOperations.SelectShape("Step Back This Shape");
            PplFeatures.StepBack();

            AssertIsSame(23, 27);
            AssertIsSame(24, 28);
            AssertIsSame(25, 29);
        }

        private void TestDrillDownUnsuccessful()
        {
            Microsoft.Office.Interop.PowerPoint.Slide slide = PpOperations.SelectSlide(34);
            slide.MoveTo(35);
            PpOperations.SelectShape("Zoom This Shape");
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to Add Animations",
                "No next slide is found. Please select the correct slide.",
                PplFeatures.DrillDown);
        }

        private void TestStepBackUnsuccessful()
        {
            Microsoft.Office.Interop.PowerPoint.Slide slide = PpOperations.SelectSlide(35);
            slide.MoveTo(1);
            PpOperations.SelectShape("Zoom This Shape");
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to Add Animations",
                "No previous slide is found. Please select the correct slide.",
                PplFeatures.StepBack);
        }


        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
