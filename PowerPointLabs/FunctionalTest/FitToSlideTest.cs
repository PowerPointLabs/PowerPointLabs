using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class FitToSlideTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "FitToSlide.pptx";
        }

        [TestMethod]
        public void FT_FitToSlideTest()
        {
            FitToWidth();
            FitToHeight();
            FitToWidthForRotatedShape();
            FitToHeightForRotatedShape();
        }

        private void FitToHeight()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShapes("pic")[1];

            PplFeatures.FitToHeight();

            var expSlide = PpOperations.SelectSlide(6);
            var expShape = PpOperations.SelectShapes("pic")[1];

            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameShape(expShape, actualShape);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void FitToWidth()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShapes("pic")[1];

            PplFeatures.FitToWidth();

            var expSlide = PpOperations.SelectSlide(5);
            var expShape = PpOperations.SelectShapes("pic")[1];

            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameShape(expShape, actualShape);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void FitToHeightForRotatedShape()
        {
            var actualSlide = PpOperations.SelectSlide(8);
            var actualShape = PpOperations.SelectShapes("pic")[1];

            PplFeatures.FitToHeight();

            var expSlide = PpOperations.SelectSlide(10);
            var expShape = PpOperations.SelectShapes("pic")[1];

            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameShape(expShape, actualShape);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void FitToWidthForRotatedShape()
        {
            var actualSlide = PpOperations.SelectSlide(8);
            var actualShape = PpOperations.SelectShapes("pic")[1];

            PplFeatures.FitToWidth();

            var expSlide = PpOperations.SelectSlide(9);
            var expShape = PpOperations.SelectShapes("pic")[1];

            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameShape(expShape, actualShape);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
