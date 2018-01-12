using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ConvertToPictureTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ShortcutsLab\\ConvertToPicture.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
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
