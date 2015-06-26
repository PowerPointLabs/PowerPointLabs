using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoZoomTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoZoom.pptx";
        }

        [TestMethod]
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

            AssertIsSame(16, 19);
            AssertIsSame(17, 20);
            AssertIsSame(18, 21);
        }

        private void TestStepBackBackground()
        {
            PplFeatures.SetZoomProperties(false, true);

            PpOperations.SelectSlide(23);
            PpOperations.SelectShape("Step Back This Shape");
            PplFeatures.StepBack();

            AssertIsSame(22, 25);
            AssertIsSame(23, 26);
            AssertIsSame(24, 27);
        }

        private void TestDrillDownUnsuccessful()
        {
            var slide = PpOperations.SelectSlide(32);
            slide.MoveTo(33);
            PpOperations.SelectShape("Zoom This Shape");
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to Add Animations",
                "No next slide is found. Please select the correct slide",
                PplFeatures.DrillDown);
        }

        private void TestStepBackUnsuccessful()
        {
            var slide = PpOperations.SelectSlide(33);
            slide.MoveTo(1);
            PpOperations.SelectShape("Zoom This Shape");
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to Add Animations",
                "No previous slide is found. Please select the correct slide",
                PplFeatures.StepBack);
        }


        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
