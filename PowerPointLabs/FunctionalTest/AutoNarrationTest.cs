using System;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AutoNarrationTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoNarrate.pptx";
        }

        [TestMethod]
        public void FT_AutoNarrationTest()
        {
            var actualSlide = PpOperations.SelectSlide(7);

            PplFeatures.AutoNarrate();

            var expSlide = PpOperations.SelectSlide(8);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }
    }
}
