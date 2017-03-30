using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class CropToSameDimTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "CropToSame.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToSameTest()
        {
            CropOneSideSuccessfully();
            CropTwoSidesSuccessfully();
            CropFourSidesSuccessfully();
            CropSmallerRefImgSuccessfully();
            CropUnevenScaledImgSuccessfully();
            CropCroppedImgSuccessfully();
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToSameNegativeTest()
        {
            CropLargerRefImgUnsuccessfully();
            CropSameSizeUnsuccessfully();
            CropOnNothingUnsuccessfully();
            CropOnShapeObjectUnsuccessfully();
        }

        #region Positive Test Cases

        public void CropOneSideSuccessfully()
        {
            CropAndCompare(4, 5);
        }

        public void CropTwoSidesSuccessfully()
        {
            CropAndCompare(7, 8);
        }

        public void CropFourSidesSuccessfully()
        {
            CropAndCompare(10, 11);
        }

        public void CropSmallerRefImgSuccessfully()
        {
            CropAndCompare(13, 14);
        }
        
        public void CropUnevenScaledImgSuccessfully()
        {
            CropAndCompare(16, 17);
        }

        public void CropCroppedImgSuccessfully()
        {
            CropAndCompare(19, 20);
        }

        public void CropAndCompare(int testSlideNo, int expectedSlideNo)
        {
            var actualSlide = PpOperations.SelectSlide(testSlideNo);
            PpOperations.SelectShapesByPrefix("selectMe");

            // Execute the Crop To Same feature
            PplFeatures.CropToSame();
            var resultShapes = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var resultShapesInPic = PpOperations.ExportSelectedShapes();

            var expSlide = PpOperations.SelectSlide(expectedSlideNo);

            var expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameLooking(expShape, expShapeInPic, expShape, expShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
        #endregion
        #region Negative Test Cases
        
        public void CropLargerRefImgUnsuccessfully()
        {
            PpOperations.SelectSlide(28);
            PpOperations.SelectShapesByPrefix("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "All selected pictures are smaller than reference shape. No cropping was done.",
                PplFeatures.CropToSame);
        }

        public void CropSameSizeUnsuccessfully()
        {
            PpOperations.SelectSlide(26);
            PpOperations.SelectShapesByPrefix("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "All selected pictures are smaller than reference shape. No cropping was done.",
                PplFeatures.CropToSame);
        }

        private void CropOnNothingUnsuccessfully()
        {
            PpOperations.SelectSlide(24);
            // don't select any shape here

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "You need to select at least 2 pictures before applying 'Crop To Same Dimensions'.",
                PplFeatures.CropToSame);
        }

        private void CropOnShapeObjectUnsuccessfully()
        {
            PpOperations.SelectSlide(22);
            PpOperations.SelectShapesByPrefix("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "'Crop To Same Dimensions' only supports picture objects.",
                PplFeatures.CropToSame);
        }

        #endregion
    }
}
