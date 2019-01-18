using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class FillSlideTest : BaseFunctionalTest
    {
        //Slide Numbers
        private const int OriginalSingleImageFillSlideSlideNo = 4;
        private const int ExpectedSingleImageFillSlideSlideNo = 5;
        private const int OriginalGroupImageFillSlideSlideNo = 7;
        private const int ExpectedGroupImageFillSlideSlideNo = 8;
        private const string ShapeToCopyPrefix = "selectMe";

        protected override string GetTestingSlideName()
        {
            return "ShortcutsLab\\FillSlide.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_FillSlideTest()
        {
            CheckFillSlide(OriginalSingleImageFillSlideSlideNo, ExpectedSingleImageFillSlideSlideNo);
            CheckFillSlide(OriginalGroupImageFillSlideSlideNo, ExpectedGroupImageFillSlideSlideNo);
        }

        private void CheckFillSlide(int originalSlideNo, int expSlideNo)
        {
            PpOperations.SelectSlide(originalSlideNo);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteToFillSlide();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private Microsoft.Office.Interop.PowerPoint.ShapeRange GetShapesByPrefix(int slideNo, string shapePrefix)
        {
            PpOperations.SelectSlide(slideNo);
            return PpOperations.SelectShapesByPrefix(shapePrefix);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(actualSlideNo);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
