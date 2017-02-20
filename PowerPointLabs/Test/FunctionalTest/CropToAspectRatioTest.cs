using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class CropToAspectRatioTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "CropToAspectRatio.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToAspectRatioTest()
        {
            CropOnePictureSuccessfully();
            CropMultiplePicturesSuccessfully();
            CropRotatedPictureSuccessfully();
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToAspectRatioNegativeTest()
        {
            CropOnNothingUnsuccessfully();
            CropOnNonPictureObjectUnsuccessfully();
        }

        #region Positive Test Cases

        public void CropOnePictureSuccessfully()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            PpOperations.SelectShape("selectMe");
            
            PplFeatures.CropToAspectRatioW1H10();

            var resultShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var resultShapeInPic = PpOperations.ExportSelectedShapes();

            var expSlide = PpOperations.SelectSlide(5);

            var expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        public void CropMultiplePicturesSuccessfully()
        {
            var actualSlide = PpOperations.SelectSlide(7);
            var shapesBeforeCrop = PpOperations.SelectShapesByPrefix("selectMe");
            Assert.AreEqual(2, shapesBeforeCrop.Count);
            
            PplFeatures.CropToAspectRatioW1H10();

            var resultShape = PpOperations.SelectShapesByPrefix("selectMe");
            var resultShapeInPic = PpOperations.ExportSelectedShapes();

            var expSlide = PpOperations.SelectSlide(8);

            var expShape = PpOperations.SelectShapesByPrefix("selectMe");
            var expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameLooking(expShape.Group(), expShapeInPic, resultShape.Group(), resultShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void CropRotatedPictureSuccessfully()
        {
            var actualSlide = PpOperations.SelectSlide(10);
            PpOperations.SelectShape("selectMe");

            PplFeatures.CropToAspectRatioW1H10();

            var resultShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var resultShapeInPic = PpOperations.ExportSelectedShapes();

            var expSlide = PpOperations.SelectSlide(11);

            var expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        #endregion
        #region Negative Test Cases

        private void CropOnNothingUnsuccessfully()
        {
            PpOperations.SelectSlide(4);
            // don't select any shape here

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "You need to select at least 1 picture before applying 'Crop To Aspect Ratio'.",
                PplFeatures.CropToAspectRatioW1H10);
        }

        private void CropOnNonPictureObjectUnsuccessfully()
        {
            PpOperations.SelectSlide(11);
            PpOperations.SelectShapesByPrefix("text");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "'Crop To Aspect Ratio' only supports picture objects.",
                PplFeatures.CropToAspectRatioW1H10);
        }

        #endregion
    }
}
