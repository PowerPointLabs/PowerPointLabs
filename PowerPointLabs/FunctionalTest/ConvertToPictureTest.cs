using FunctionalTest.util;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class ConvertToPictureTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ConvertToPicture.pptx";
        }

        [TestMethod]
        public void FT_ConvertToShapeTest()
        {
            ConvertSingleObjToPicture();
            CovertGroupObjToPicture();
        }

        private void CovertGroupObjToPicture()
        {
            var actualSlide = PpOperations.SelectSlide(7);
            PpOperations.SelectShapes("pic");

            PplFeatures.ConvertToPic();

            var sh = PpOperations.SelectShapes("Picture 1")[1];
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            var expSlide = PpOperations.SelectSlide(8);
            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void ConvertSingleObjToPicture()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            PpOperations.SelectShapes("pic");

            PplFeatures.ConvertToPic();

            var sh = PpOperations.SelectShapes("Picture 1")[1];
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            var expSlide = PpOperations.SelectSlide(5);
            PpOperations.SelectShapes("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
