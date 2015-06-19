using System.Deployment.Internal;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class ZoomToAreaTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ZoomToArea.pptx";
        }

        [TestMethod]
        public void FT_ZoomToAreaTest()
        {
            // Do tests in reverse order because added slides change slide numbers lower down.
            TestMultipleZoom();
            TestSingleZoom();
        }

        private void TestMultipleZoom()
        {
            PpOperations.SelectSlide(10);
            PpOperations.SelectShapes(new[] { "First ZoomShape", "Second ZoomShape", "Third ZoomShape", "Fourth ZoomShape" });
            PplFeatures.AddZoomToArea();

            for (int i = 0; i < 10; ++i)
            {
                AssertIsSame(10 + i, 20 + i);
            }
        }

        private void TestSingleZoom()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes("First ZoomShape");
            PplFeatures.AddZoomToArea();

            for (int i = 0; i < 4; ++i)
            {
                AssertIsSame(4 + i, 8 + i);
            }
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
