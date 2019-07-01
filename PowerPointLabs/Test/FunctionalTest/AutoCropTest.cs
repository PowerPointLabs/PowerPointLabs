using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoCropTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "CropLab\\AutoCrop.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoCropTest()
        {
            CropOneShapeSuccessfully();
            CropMultipleShapesSuccessfully();
            CropRotatedShapeSuccessfully();
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoCropNegativeTest()
        {
            CropOnNothingUnsuccessfully();
            CropOnPictureObjectUnsuccessfully();
        }

        #region Positive Test Cases

        public void CropOneShapeSuccessfully()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(4);
            PpOperations.SelectShape("selectMe");

            // Execute the Crop To Shape feature
            PplFeatures.AutoCrop();

            Microsoft.Office.Interop.PowerPoint.Shape resultShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            System.IO.FileInfo resultShapeInPic = PpOperations.ExportSelectedShapes();

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(5);

            Microsoft.Office.Interop.PowerPoint.Shape expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            System.IO.FileInfo expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").SafeDelete();

            SlideUtil.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        public void CropMultipleShapesSuccessfully()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapesBeforeCrop = PpOperations.SelectShapesByPrefix("selectMe");
            Assert.AreEqual(6, shapesBeforeCrop.Count);

            // Execute the Crop To Shape feature
            PplFeatures.AutoCrop();

            // the result shape after crop multiple shapes will have name starts with
            // Group
            Microsoft.Office.Interop.PowerPoint.Shape resultShape = PpOperations.SelectShapesByPrefix("Group")[1];
            System.IO.FileInfo resultShapeInPic = PpOperations.ExportSelectedShapes();

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);

            Microsoft.Office.Interop.PowerPoint.Shape expShape = PpOperations.SelectShapesByPrefix("Group")[1];
            System.IO.FileInfo expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").SafeDelete();

            SlideUtil.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void CropRotatedShapeSuccessfully()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(10);
            PpOperations.SelectShape("selectMe");

            // Execute the Crop To Shape feature
            PplFeatures.AutoCrop();

            Microsoft.Office.Interop.PowerPoint.Shape resultShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            System.IO.FileInfo resultShapeInPic = PpOperations.ExportSelectedShapes();

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(11);

            Microsoft.Office.Interop.PowerPoint.Shape expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            System.IO.FileInfo expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").SafeDelete();

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
                "You need to select at least 1 shape before applying 'Crop To Shape'.",
                PplFeatures.AutoCrop);
        }

        private void CropOnPictureObjectUnsuccessfully()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes(new List<string> {"selectMe", "pic"});

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Error", 
                "'Crop To Shape' only supports shape objects.",
                PplFeatures.AutoCrop);
        }

        #endregion
    }
}
