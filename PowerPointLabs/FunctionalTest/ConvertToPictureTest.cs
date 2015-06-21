using FunctionalTest.util;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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
        public void FT_ConvertToPictureTest()
        {
            ConvertSingleObjToPicture();
            CovertGroupObjToPicture();
        }

        private void CovertGroupObjToPicture()
        {
            var actualSlide = PpOperations.SelectSlide(7);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            var sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            var expSlide = PpOperations.SelectSlide(8);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void ConvertSingleObjToPicture()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            var sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            var expSlide = PpOperations.SelectSlide(5);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
