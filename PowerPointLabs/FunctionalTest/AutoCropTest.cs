using FunctionalTest.util;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoCropTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoCrop.pptx";
        }

        [TestMethod]
        public void FT_AutoCropTest()
        {
            CropOneShapeSuccessfully();
            CropMultipleShapesSuccessfully();
        }

        public void CropOneShapeSuccessfully()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            var shapeBeforeChange = PpOperations.SelectShapes("selectMe")[1];
            Assert.AreEqual("selectMe", shapeBeforeChange.Name);

            // Execute the Crop To Shape feature
            PplFeatures.AutoCrop();
            var resultShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var resultShapeInPic = PpOperations.ExportSelectedShapes();
            Assert.IsTrue(resultShape.Name.StartsWith("selectMe"));

            var expSlide = PpOperations.SelectSlide(5);
            var expShape = PpOperations.SelectShapesByPrefix("selectMe")[1];
            var expShapeInPic = PpOperations.ExportSelectedShapes();
            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideComparer.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideComparer.IsSameLooking(expSlide, actualSlide);
        }

        public void CropMultipleShapesSuccessfully()
        {
            var actualSlide = PpOperations.SelectSlide(7);
            var shapesBeforeCrop = PpOperations.SelectShapesByPrefix("selectMe");
            Assert.AreEqual(6, shapesBeforeCrop.Count);

            // Execute the Crop To Shape feature
            PplFeatures.AutoCrop();

            // the result shape after crop multiple shapes will have name starts with
            // Group
            var resultShape = PpOperations.SelectShapesByPrefix("Group")[1];
            var resultShapeInPic = PpOperations.ExportSelectedShapes();

            var expSlide = PpOperations.SelectSlide(8);

            var expShape = PpOperations.SelectShapesByPrefix("Group")[1];
            var expShapeInPic = PpOperations.ExportSelectedShapes();

            // remove elements that affect comparing slides
            // e.g. "Expected" textbox
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideComparer.IsSameLooking(expShape, expShapeInPic, resultShape, resultShapeInPic);
            SlideComparer.IsSameLooking(expSlide, actualSlide);
        }
    }
}
