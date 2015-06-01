using System;
using System.Drawing;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoCropTest : BaseFunctionalTest
    {
        [TestMethod]
        public void TestAutoCrop()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            var shapeBeforeChange = PpOperations.SelectShapes("selectMe")[1];
            Assert.AreEqual("selectMe", shapeBeforeChange.Name);
            PplFeatures.AutoCrop();
            var range = PpOperations.SelectShapesByPrefix("selectMe");
            var resultShape = range[1];
            Assert.IsTrue(resultShape.Name.StartsWith("selectMe"));

            var expSlide = PpOperations.SelectSlide(5);
            var text = PpOperations.SelectShapesByPrefix("text");
            text.Delete();

            actualSlide.Export(Path.GetTempPath() + "1.png", "PNG");
            expSlide.Export(Path.GetTempPath() + "2.png", "PNG");

            var b1 = new Bitmap(Path.GetTempPath() + "1.png");
            var b2 = new Bitmap(Path.GetTempPath() + "2.png");
            var resultOfAutoCrop = ImageComparer.Compare(b1, b2);
            Assert.AreEqual(ImageComparer.CompareResult.CompareOk, resultOfAutoCrop);
        }

        protected override string GetSlideName()
        {
            return "AutoCrop.pptx";
        }
    }
}
