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

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }
    }
}
