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
            TestStepBack();
            TestDrillDown();
            TestDrillDownUnsuccessful();
            TestStepBackUnsuccessful();
        }

        private void TestDrillDown()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShape("Drill Down This Shape");
            PplFeatures.DrillDown();

            AssertIsSame(4, 7);
            AssertIsSame(5, 8);
            AssertIsSame(6, 9);
        }

        private void TestStepBack()
        {
            PpOperations.SelectSlide(11);
            PpOperations.SelectShape("Step Back This Shape");
            PplFeatures.StepBack();

            AssertIsSame(10, 13);
            AssertIsSame(11, 14);
            AssertIsSame(12, 15);
        }

        private void TestDrillDownUnsuccessful()
        {
            var slide = PpOperations.SelectSlide(18);
            slide.MoveTo(19);
            PpOperations.SelectShape("Zoom This Shape");
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to Add Animations",
                "No next slide is found. Please select the correct slide",
                PplFeatures.DrillDown);
        }

        private void TestStepBackUnsuccessful()
        {
            var slide = PpOperations.SelectSlide(19);
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
