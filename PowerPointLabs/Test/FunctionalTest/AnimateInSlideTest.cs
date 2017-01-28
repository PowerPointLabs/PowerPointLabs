using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AnimateInSlideTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AnimateInSlide.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AnimateInSlideTest()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes(new List<string> { "Rectangle 2", "Rectangle 5", "Rectangle 6" });

            PplFeatures.AnimateInSlide();

            var actualSlide = PpOperations.SelectSlide(4);
            var expSlide = PpOperations.SelectSlide(5);

            // remove text "Expected"
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AnimateInSlideStraightLinesTest()
        {
            PpOperations.SelectSlide(21);
            PpOperations.SelectShapes(new List<string> { "Straight Arrow Connector 61",
                                                         "Straight Arrow Connector 63",
                                                         "Straight Arrow Connector 66" });

            PplFeatures.AnimateInSlide();

            var actualSlide = PpOperations.SelectSlide(21);
            var expSlide = PpOperations.SelectSlide(22);

            // remove text "Expected"
            PpOperations.SelectShape("Text Label Expected Output")[1].Delete();
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        public void FT_AnimateInSlideFlippedTest()
        {
            PpOperations.SelectSlide(10);

            PpOperations.SelectShapes(new List<string> { "Arrow 1a", "Arrow 1b" });
            PplFeatures.AnimateInSlide();

            PpOperations.SelectShapes(new List<string> { "Arrow 2a", "Arrow 2b" });
            PplFeatures.AnimateInSlide();

            PpOperations.SelectShapes(new List<string> { "Bolt 1a", "Bolt 1b" });
            PplFeatures.AnimateInSlide();

            var actualSlide = PpOperations.SelectSlide(10);
            var expSlide = PpOperations.SelectSlide(11);

            // remove text "Expected"
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}