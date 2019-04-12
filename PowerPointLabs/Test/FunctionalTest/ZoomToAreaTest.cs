using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ZoomToAreaTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ZoomLab\\ZoomToArea.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_ZoomToAreaTest()
        {
            // Do tests in reverse order because added slides change slide numbers lower down.
            TestMultipleZoomConfig4();
            TestMultipleZoomConfig3();
            TestMultipleZoomConfig2();
            TestMultipleZoomConfig1();
            TestSingleZoomConfig2();
            TestSingleZoomConfig1();
        }

        private void TestMultipleZoomConfig4()
        {
            PplFeatures.SetZoomProperties(false, false);

            PpOperations.SelectSlide(42);
            PpOperations.SelectShapes(new[] { "First ZoomShape", "Second ZoomShape", "Third ZoomShape", "Fourth ZoomShape" });
            PplFeatures.AddZoomToArea();

            AssertAreSame(42, 44, 2);
        }

        private void TestMultipleZoomConfig3()
        {
            PplFeatures.SetZoomProperties(false, true);

            PpOperations.SelectSlide(30);
            PpOperations.SelectShapes(new[] { "First ZoomShape", "Second ZoomShape", "Third ZoomShape", "Fourth ZoomShape" });
            PplFeatures.AddZoomToArea();

            AssertAreSame(30, 40, 10);
        }

        private void TestMultipleZoomConfig2()
        {
            PplFeatures.SetZoomProperties(true, false);

            PpOperations.SelectSlide(26);
            PpOperations.SelectShapes(new[] { "First ZoomShape", "Second ZoomShape", "Third ZoomShape", "Fourth ZoomShape" });
            PplFeatures.AddZoomToArea();

            AssertAreSame(26, 28, 2);
        }

        private void TestMultipleZoomConfig1()
        {
            PplFeatures.SetZoomProperties(true, true);

            PpOperations.SelectSlide(14);
            PpOperations.SelectShapes(new[] { "First ZoomShape", "Second ZoomShape", "Third ZoomShape", "Fourth ZoomShape" });
            PplFeatures.AddZoomToArea();

            AssertAreSame(14, 24, 10);
        }

        private void TestSingleZoomConfig2()
        {
            PplFeatures.SetZoomProperties(false, false);

            PpOperations.SelectSlide(10);
            PpOperations.SelectShape("First ZoomShape");
            PplFeatures.AddZoomToArea();

            AssertAreSame(10, 12, 2);
        }

        private void TestSingleZoomConfig1()
        {
            PplFeatures.SetZoomProperties(true, true);

            PpOperations.SelectSlide(4);
            PpOperations.SelectShape("First ZoomShape");
            PplFeatures.AddZoomToArea();

            AssertAreSame(4, 8, 4);
        }

        private void AssertAreSame(int actualStartSlideIndex, int expectedStartSlideIndex, int slideCount)
        {
            for (int i = 0; i < slideCount; ++i)
            {
                AssertIsSame(actualStartSlideIndex + i, expectedStartSlideIndex + i);
            }
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            // Similarity tolerance adjusted because different machines generate slightly different slides
            SlideUtil.IsSameLooking(expectedSlide, actualSlide, 0.90);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
