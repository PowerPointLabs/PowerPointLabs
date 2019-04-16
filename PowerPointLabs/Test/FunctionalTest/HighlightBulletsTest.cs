using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class HighlightBulletsTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "HighlightLab\\HighlightPoints.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_HighlightBulletsTest()
        {
            // Do tests in reverse order because added slides change slide numbers lower down.
            TestHighlightPoints_SelectEndOfText();
            TestHighlightBackground_SelectText();
            TestHighlightBackground_SelectTextBoxes();
            TestHighlightBackground_SelectSlide();
            TestHighlightPoints_SelectText();
            TestHighlightPoints_SelectTextBoxes();
            TestHighlightPoints_SelectSlide();
        }

        private void TestHighlightPoints_SelectEndOfText()
        {
            PpOperations.SelectSlide(22);
            PpOperations.SelectTextInShape("First Textbox", 414, 414);
            PplFeatures.HighlightPoints();

            AssertIsSame(22, 23);
        }

        private void TestHighlightBackground_SelectText()
        {
            PpOperations.SelectSlide(19);
            PpOperations.SelectAllTextInShape("First Textbox");
            PplFeatures.HighlightBackground();

            AssertIsSame(19, 20);
        }

        private void TestHighlightBackground_SelectTextBoxes()
        {
            PpOperations.SelectSlide(16);
            PpOperations.SelectShapes(new[] { "First Textbox", "Second Textbox", "Third TextBox" });
            PplFeatures.HighlightBackground();

            AssertIsSame(16, 17);
        }

        private void TestHighlightBackground_SelectSlide()
        {
            PpOperations.SelectSlide(13);
            PplFeatures.HighlightBackground();

            AssertIsSame(13, 14);
        }

        private void TestHighlightPoints_SelectText()
        {
            PpOperations.SelectSlide(10);
            PpOperations.SelectAllTextInShape("First Textbox");
            PplFeatures.HighlightPoints();

            AssertIsSame(10, 11);
        }

        private void TestHighlightPoints_SelectTextBoxes()
        {
            PpOperations.SelectSlide(7);
            PpOperations.SelectShapes(new[] { "First Textbox", "Second Textbox", "Third TextBox" });
            PplFeatures.HighlightPoints();

            AssertIsSame(7, 8);
        }

        private void TestHighlightPoints_SelectSlide()
        {
            PpOperations.SelectSlide(4);
            PplFeatures.HighlightPoints();

            AssertIsSame(4, 5);
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
