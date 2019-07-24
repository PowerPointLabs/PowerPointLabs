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
            ConvertGroupObjToPicture();
            ConvertSingleObjInGroupToPicutre();
            IsClipboardRestored();
        }

        private void ConvertSingleObjInGroupToPicutre()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(13);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            Shape sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(14);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void ConvertGroupObjToPicture()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            Shape sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void ConvertSingleObjToPicture()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(4);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            Shape sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(5);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void IsClipboardRestored()
        {
            CheckIfClipboardIsRestored(() =>
            {
                PpOperations.SelectShape("pic");
                PplFeatures.ConvertToPic();
            }, 10, "pictocopy", 11, "text 3", "copied");
        }
    }
}
