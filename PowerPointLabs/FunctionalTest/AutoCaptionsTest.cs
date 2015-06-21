using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoCaptionsTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoCaptions.pptx";
        }

        [TestMethod]
        public void FT_AutoCaptionsTest()
        {
            var actualSlide = PpOperations.SelectSlide(4);
            ThreadUtil.WaitFor(1000);

            PplFeatures.AutoCaptions();

            var expSlide = PpOperations.SelectSlide(5);
            PpOperations.SelectShape("text 3").Delete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
