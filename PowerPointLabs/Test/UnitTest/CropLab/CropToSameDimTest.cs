using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.CropLab;

namespace Test.UnitTest.CropLab
{
    [TestClass]
    public class CropToSameDimTest : BaseCropLabTest
    {
        private List<string> selectMultipleShapesNames = new List<string> { "selectMe2", "selectMe1" };

        protected override string GetTestingSlideName()
        {
            return "CropLab\\CropToSame.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropToSameTest()
        {
            CropOneSideSuccessfully();
            CropTwoSidesSuccessfully();
            CropFourSidesSuccessfully();
            CropSmallerRefImgSuccessfully();
            CropUnevenScaledImgSuccessfully();
            CropCroppedImgSuccessfully();
            CropCustomAnchorSuccessfully();
            CropWithShapeAsReferenceSuccessfully();
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

        public void CropCustomAnchorSuccessfully()
        {
            CropLabSettings.AnchorPosition = AnchorPosition.BottomRight;
            CropAndCompare(22, 23);
            CropLabSettings.AnchorPosition = AnchorPosition.Reference;
        }

        public void CropWithShapeAsReferenceSuccessfully()
        {
            CropAndCompare(33, 34);
        }

        public void CropAndCompare(int testSlideNo, int expectedSlideNo)
        {
            // Execute the Crop To Same feature
            var testShapes = GetShapes(testSlideNo, selectMultipleShapesNames);
            var expShapes = GetShapes(expectedSlideNo, selectMultipleShapesNames);
            CropToSame.CropSelection(testShapes);

            testShapes = GetShapes(testSlideNo, selectMultipleShapesNames);
            CheckShapes(testShapes, expShapes);
            
        }
        #endregion
            
    }
}
