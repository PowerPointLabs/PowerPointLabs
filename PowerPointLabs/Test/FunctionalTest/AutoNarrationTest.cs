using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoNarrationTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "NarrationsLab\\AutoNarrate.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoNarrationTest()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);

            PplFeatures.AutoNarrate();
            ThreadUtil.WaitFor(1000); // need to wait for loading
            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);
            ThreadUtil.WaitFor(1000); // need to wait for loading
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }
    }
}
