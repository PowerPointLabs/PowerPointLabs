﻿using System.Collections.Generic;

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

        private List<string> selectOneShapeNames = new List<string> { "selectMe" };
        private List<string> selectMultipleShapesNames = new List<string> { "selectMe1", "selectMe2" };

        protected override string GetTestingSlideName()
        {
            return "CropOutPadding.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingOnePicture()
        {
            var actualShapes = GetShapes(SlideNumberOnePictureActual, selectOneShapeNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            var expectedShapes = GetShapes(SlideNumberOnePictureExpected, selectOneShapeNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingMultiplePictures()
        {
            var actualShapes = GetShapes(SlideNumberMultiplePicturesActual, selectMultipleShapesNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            var expectedShapes = GetShapes(SlideNumberMultiplePicturesExpected, selectMultipleShapesNames);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CropOutPaddingRotatedPicture()
        {
            var actualShapes = GetShapes(SlideNumberRotatedPictureActual, selectOneShapeNames);
            actualShapes = CropOutPadding.Crop(actualShapes);

            var expectedShapes = GetShapes(SlideNumberRotatedPictureExpected, selectOneShapeNames);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
