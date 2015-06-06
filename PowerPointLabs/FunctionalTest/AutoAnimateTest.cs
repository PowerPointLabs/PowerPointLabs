using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoAnimateTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoAnimate.pptx";
        }

        [TestMethod]
        public void FT_AutoAnimateTest()
        {
            AutoAnimateSuccessfully();
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

            SlideComparer.IsSameAnimations(expSlide, actualSlide);
            SlideComparer.IsSameLooking(expSlide, actualSlide);
        }
    }
}
