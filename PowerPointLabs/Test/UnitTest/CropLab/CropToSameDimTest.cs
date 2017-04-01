using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.CropLab;

namespace Test.UnitTest.CropLab
{
    [TestClass]
    public class CropToSameDimTest : BaseCropLabTest
    {
        private const int SlideNumberOnePictureActual = 4;
        private const int SlideNumberOnePictureExpected = 5;
        private const int SlideNumberMultiplePicturesActual = 7;
        private const int SlideNumberMultiplePicturesExpected = 8;
        private const int SlideNumberRotatedPictureActual = 10;
        private const int SlideNumberRotatedPictureExpected = 11;

        private List<string> selectOneShapeNames = new List<string> { "selectMe" };
        private List<string> selectMultipleShapesNames = new List<string> { "selectMe1", "selectMe2" };

        protected override string GetTestingSlideName()
        {
            return "CropToSame.pptx";
        }



        [TestMethod]
        [TestCategory("UT")]
        public void UT_CropToSameTest()
        {
            CropOneSideSuccessfully();
            CropTwoSidesSuccessfully();
            CropFourSidesSuccessfully();
            CropSmallerRefImgSuccessfully();
            CropUnevenScaledImgSuccessfully();
            CropCroppedImgSuccessfully();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void UT_CropToSameNegativeTest()
        {
            /*
            CropLargerRefImgUnsuccessfully();
            CropSameSizeUnsuccessfully();
            CropOnNothingUnsuccessfully();
            CropOnShapeObjectUnsuccessfully();
            */
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
            // Execute the Crop To Same feature
            var testShapes = GetShapes(testSlideNo, selectMultipleShapesNames);
            var expShapes = GetShapes(expectedSlideNo, selectMultipleShapesNames);
            CropToSame.CropSelection(testShapes);
            CheckShapes(testShapes, expShapes);
            
        }
        #endregion
        #region Negative Test Cases
        /*
        public void CropLargerRefImgUnsuccessfully()
        {
            PpOperations.SelectSlide(28);
            PpOperations.SelectShapesByPrefix("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "Target picture dimensions are equal to or smaller than reference shape. No cropping was done.",
                PplFeatures.CropToSame);
        }

        public void CropSameSizeUnsuccessfully()
        {
            PpOperations.SelectSlide(26);
            PpOperations.SelectShapesByPrefix("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "Target picture dimensions are equal to or smaller than reference shape. No cropping was done.",
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
        */
        #endregion
            
    }
}
