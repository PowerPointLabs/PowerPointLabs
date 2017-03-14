﻿using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class CropToSlideTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "CropToSlide.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToSlideTest()
        {
            CropOnePicInSlideSuccessfully();
            CropOnePicOneEdgeSuccessfully();
            CropOnePicMultipleEdgesSuccessfully();
            CropOneRotatedPicOneEdgeSuccessfully();
            CropOneRotatedPicMultipleEdgesSuccessfully();
            CropMultiplePicsSuccessfully();
            CropMultipleRotatedPicsSuccessfully();
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CropToSlideNegativeTest()
        {
            CropOnNothingUnsuccessfully();
            CropOnShapeObjectUnsuccessfully();
        }

        #region Positive Test Cases

        public void CropOnePicInSlideSuccessfully()
        {
            CropAndCompare(4, 5);
        }

        public void CropOnePicOneEdgeSuccessfully()
        {
            CropAndCompare(7, 8);
        }

        public void CropOnePicMultipleEdgesSuccessfully()
        {
            CropAndCompare(10, 11);
        }

        public void CropOneRotatedPicOneEdgeSuccessfully()
        {
            CropAndCompare(13, 14);
        }

        public void CropOneRotatedPicMultipleEdgesSuccessfully()
        {
            CropAndCompare(16, 17);
        }

        public void CropMultiplePicsSuccessfully()
        {
            CropAndCompare(19, 20);
        }

        public void CropMultipleRotatedPicsSuccessfully()
        {
            CropAndCompare(22, 23);
        }

        public void CropAndCompare(int testSlideNo, int expectedSlideNo)
        {
            var actualSlide = PpOperations.SelectSlide(testSlideNo);
            PpOperations.SelectShapesByPrefix("selectMe");

            // Execute the Crop To Slide feature
            PplFeatures.CropToSlide();
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

        private void CropOnNothingUnsuccessfully()
        {
            PpOperations.SelectSlide(27);
            // don't select any shape here

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "You need to select at least 1 picture before applying 'Crop To Slide'.",
                PplFeatures.CropToSlide);
        }

        private void CropOnShapeObjectUnsuccessfully()
        {
            PpOperations.SelectSlide(25);
            PpOperations.SelectShape("selectMe");

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error",
                "'Crop To Slide' only supports picture objects.",
                PplFeatures.CropToSlide);
        }
        
        #endregion
    }
}
