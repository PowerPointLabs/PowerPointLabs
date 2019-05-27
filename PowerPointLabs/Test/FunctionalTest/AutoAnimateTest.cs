using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoAnimateTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AnimationLab\\AutoAnimate.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoAnimateSuccessfully()
        {
            AutoAnimateSuccessfully();
        }

        // create a new test, since the previous one will change the later's slide index..
        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoAnimateWithCopyPasteSuccessfully()
        {
            AutoAnimateWithCopyPasteShapesSuccessfully();
        }
        
        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoAnimateStraightLinesSuccessfully()
        {
            AutoAnimateStraightLinesSuccessfully();
        }

        private static void AutoAnimateSuccessfully()
        {
            PpOperations.SelectSlide(4);

            PplFeatures.AutoAnimate();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(5);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(7);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void AutoAnimateWithCopyPasteShapesSuccessfully()
        {
            PpOperations.SelectSlide(8);
            ShapeRange pastingShapes = PpOperations.SelectShapes(new List<string> {"Notched Right Arrow", "Group"});
            PPLClipboard.Instance.LockClipboard();
            pastingShapes.Copy();

            Slide targetSlide = PpOperations.SelectSlide(9);
            targetSlide.Shapes.Paste();
            PPLClipboard.Instance.ReleaseClipboard();

            Assert.IsNotNull(PpOperations.SelectShape("Notched Right Arrow"), 
                "Copy-Paste failed, this task is flaky so please re-run.");
            Shape sh1 = PpOperations.SelectShape("Notched Right Arrow")[1];
            sh1.Rotation += 90;
            Shape sh2 = PpOperations.SelectShape("Group")[1];
            sh2.Rotation += 90;

            // go back to slide 8
            PpOperations.SelectSlide(8);

            PplFeatures.AutoAnimate();

            Slide actualSlide = PpOperations.SelectSlide(9);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            Slide expSlide = PpOperations.SelectSlide(11);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            // TODO: actually this expected slide looks a bit strange..
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void AutoAnimateStraightLinesSuccessfully()
        {
            PpOperations.SelectSlide(38);

            PplFeatures.AutoAnimate();

            Slide actualSlide = PpOperations.SelectSlide(39);
            // remove elements that affect comparing slides
            PpOperations.SelectShape("Text Label Initial Slide")[1].Delete();

            Slide expSlide = PpOperations.SelectSlide(41);
            // remove elements that affect comparing slides
            PpOperations.SelectShape("Text Label Expected Output")[1].Delete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
