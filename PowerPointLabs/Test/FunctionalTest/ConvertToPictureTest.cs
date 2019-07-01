using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
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
            IsClipboardRestored();
        }

        private void CovertGroupObjToPicture()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);
            PpOperations.SelectShape("pic");

            PplFeatures.ConvertToPic();

            Shape sh = PpOperations.SelectShapesByPrefix("Picture")[1] as Shape;
            Assert.AreEqual(MsoShapeType.msoPicture, sh.Type);

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);
            PpOperations.SelectShape("text 3")[1].SafeDelete();
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
            PpOperations.SelectShape("text 3")[1].SafeDelete();
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
