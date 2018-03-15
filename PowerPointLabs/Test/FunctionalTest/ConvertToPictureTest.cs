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
            CheckIfClipboardIsRestored();
        }

        private void CovertGroupObjToPicture()
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

        private static void CheckIfClipboardIsRestored()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(10);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapeToBeCopied = PpOperations.SelectShape("pictocopy");
            Assert.AreEqual(1, shapeToBeCopied.Count);
            // Add "pictocopy" to clipboard
            shapeToBeCopied.Copy();

            // Normally run convert to pic function
            PpOperations.SelectShape("pic");
            PplFeatures.ConvertToPic();

            // Paste whatever in clipboard
            Microsoft.Office.Interop.PowerPoint.ShapeRange newShape = actualSlide.Shapes.Paste();

            // Check if pasted shape is also called "pictocopy"
            Assert.AreEqual("pictocopy", newShape.Name);

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(11);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
