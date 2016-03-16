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
    }
}
