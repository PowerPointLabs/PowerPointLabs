using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.CropLab;

namespace Test.UnitTest.CropLab
{
    [TestClass]
    public class CropOutPaddingTest : BaseCropLabTest
    {
        private const int SlideNumberOnePictureActual = 4;
        private const int SlideNumberOnePictureExpected = 5;
        private const int SlideNumberMultiplePicturesActual = 7;
        private const int SlideNumberMultiplePicturesExpected = 8;
        private const int SlideNumberRotatedPictureActual = 10;
        private const int SlideNumberRotatedPictureExpected = 11;
        private const int SlideNumberOneChildPictureActual = 13;
        private const int SlideNumberOneChildPictureExpected = 14;
        private const int SlideNumberMultipleChildPicturesActual = 16;
        private const int SlideNumberMultipleChildPicturesExpected = 17;

        private List<string> selectOneShapeNames = new List<string> { "selectMe" };
        private List<string> selectMultipleShapesNames = new List<string> { "selectMe1", "selectMe2" };

        protected override string GetTestingSlideName()
        {
            return "CropLab\\CropOutPadding.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingOnePicture()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNumberOnePictureActual, selectOneShapeNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNumberOnePictureExpected, selectOneShapeNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingMultiplePictures()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNumberMultiplePicturesActual, selectMultipleShapesNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNumberMultiplePicturesExpected, selectMultipleShapesNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingRotatedPicture()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNumberRotatedPictureActual, selectOneShapeNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNumberRotatedPictureExpected, selectOneShapeNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingOneChildPicture()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNumberOneChildPictureActual, selectOneShapeNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNumberOneChildPictureExpected, selectOneShapeNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingMultipleChildPictures()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNumberMultipleChildPicturesActual, selectMultipleShapesNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNumberMultipleChildPicturesExpected, selectMultipleShapesNames);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
