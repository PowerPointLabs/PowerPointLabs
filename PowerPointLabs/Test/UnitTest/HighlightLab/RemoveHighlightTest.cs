using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.HighlightLab;
using PowerPointLabs.Models;

using Test.Util;

namespace Test.UnitTest.HighlightLab
{
    [TestClass]
    public class RemoveHighlightTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "HighlightLab\\HighlightPoints.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void RemoveHighlightingTest()
        {
            TestRemoveHighlighting_HighlightText();
            TestRemoveHighlighting_HighlightBackground();
            TestRemoveHighlighting_HighlightPoints();
        }

        private void TestRemoveHighlighting_HighlightText()
        {
            RemoveHighlightAndCompare(31, 32);
        }

        private void TestRemoveHighlighting_HighlightBackground()
        {
            RemoveHighlightAndCompare(28, 29);
        }

        private void TestRemoveHighlighting_HighlightPoints()
        {
            RemoveHighlightAndCompare(25, 26);
        }
        
        private void RemoveHighlightAndCompare(int testSlideNo, int expectedSlideNo)
        {
            PpOperations.SelectSlide(testSlideNo);
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());
            RemoveHighlighting.RemoveHighlight(currentSlide);
            AssertIsSame(testSlideNo, expectedSlideNo);
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
