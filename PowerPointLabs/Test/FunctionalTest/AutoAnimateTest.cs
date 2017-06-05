using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoAnimateTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoAnimate.pptx";
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
        public void FT_Flaky_AutoAnimateWithCopyPasteSuccessfully()
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

            var actualSlide = PpOperations.SelectSlide(5);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            var expSlide = PpOperations.SelectSlide(7);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private static void AutoAnimateWithCopyPasteShapesSuccessfully()
        {
            PpOperations.SelectSlide(8);
            PpOperations.SelectShapes(new List<string> {"Notched Right Arrow 3", "Group 2"});
            // use keyboard to copy & paste,
            // otherwise API's copy & paste won't trigger special clipboard event.
            KeyboardUtil.Copy();

            PpOperations.SelectSlide(9);
            KeyboardUtil.Paste();

            Assert.IsNotNull(PpOperations.SelectShape("Notched Right Arrow 3"), 
                "Copy-Paste failed, this task is flaky so please re-run.");
            var sh1 = PpOperations.SelectShape("Notched Right Arrow 3")[1];
            sh1.Rotation += 90;
            var sh2 = PpOperations.SelectShape("Group 2")[1];
            sh2.Rotation += 90;

            // go back to slide 8
            PpOperations.SelectSlide(8);

            PplFeatures.AutoAnimate();

            var actualSlide = PpOperations.SelectSlide(9);
            // remove elements that affect comparing slides
            PpOperations.SelectShapesByPrefix("text").Delete();

            var expSlide = PpOperations.SelectSlide(11);
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

            var actualSlide = PpOperations.SelectSlide(39);
            // remove elements that affect comparing slides
            PpOperations.SelectShape("Text Label Initial Slide")[1].Delete();

            var expSlide = PpOperations.SelectSlide(41);
            // remove elements that affect comparing slides
            PpOperations.SelectShape("Text Label Expected Output")[1].Delete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
