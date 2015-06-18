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
        }

        private void TestDrillDown()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes("Drill Down This Shape");
            PplFeatures.DrillDown();

            AssertIsSame(4, 7);
            AssertIsSame(5, 8);
            AssertIsSame(6, 9);
        }

        private void TestStepBack()
        {
            PpOperations.SelectSlide(11);
            PpOperations.SelectShapes("Step Back This Shape");
            PplFeatures.StepBack();

            AssertIsSame(10, 13);
            AssertIsSame(11, 14);
            AssertIsSame(12, 15);
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
