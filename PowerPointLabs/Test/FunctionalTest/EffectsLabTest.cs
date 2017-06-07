using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class EffectsLabTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "EffectsLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_EffectsLabTest()
        {
            PplFeatures.BlurrinessOverlay("EffectsLabBlurBackground", true);
            TestRemainderEffect(52, PplFeatures.SepiaBackgroundEffect);
            TestRemainderEffect(49, PplFeatures.BlurBackgroundEffect);
            TestRemainderEffect(46, PplFeatures.BlurBackgroundEffect);
            PplFeatures.BlurrinessOverlay("EffectsLabBlurRemainder", true);
            TestRemainderEffect(43, PplFeatures.BlurRemainderEffect);
            TestRemainderEffect(40, PplFeatures.BlurRemainderEffect);
            TestRemainderEffect(37, PplFeatures.SepiaBackgroundEffect);
            PplFeatures.BlurrinessOverlay("EffectsLabBlurBackground", false);
            TestRemainderEffect(34, PplFeatures.BlurBackgroundEffect);
            TestRemainderEffect(31, PplFeatures.SepiaRemainderEffect);
            TestRemainderEffect(28, PplFeatures.SepiaRemainderEffect);
            TestRemainderEffect(25, PplFeatures.GothamRemainderEffect);
            TestRemainderEffect(22, PplFeatures.BlackAndWhiteBackgroundEffect);
            TestRemainderEffect(19, PplFeatures.GreyScaleRemainderEffect);
            PplFeatures.BlurrinessOverlay("EffectsLabBlurRemainder", false);
            TestRemainderEffect(16, PplFeatures.BlurRemainderEffect);
            TestEffect(14, PplFeatures.BlurSelectedEffect);
            TestEffect(12, PplFeatures.BlurSelectedEffect);
            PplFeatures.BlurrinessOverlay("EffectsLabBlurSelected", true);
            TestEffect(10, PplFeatures.BlurSelectedEffect);
            TestEffect(8, PplFeatures.BlurSelectedEffect);
            TestEffect(6, PplFeatures.MagnifyingGlassEffect);
            TestEffect(4, PplFeatures.TransparentEffect);
        }

        private void TestRemainderEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            AssertIsSame(startIdx, startIdx + 2);
            AssertIsSame(startIdx + 1, startIdx + 3);
        }

        private void TestEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            AssertIsSame(startIdx, startIdx + 1);
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
