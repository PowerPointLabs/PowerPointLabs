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
            TestFrostedGlass();
            TestRemainderEffect(26, PplFeatures.SepiaBackgroundEffect);
            TestRemainderEffect(23, PplFeatures.BlurBackgroundEffect);
            TestRemainderEffect(20, PplFeatures.SepiaRemainderEffect);
            TestRemainderEffect(17, PplFeatures.GothamRemainderEffect);
            TestRemainderEffect(14, PplFeatures.BlackAndWhiteBackgroundEffect);
            TestRemainderEffect(11, PplFeatures.GreyScaleRemainderEffect);
            TestRemainderEffect(8, PplFeatures.BlurRemainderEffect);
            TestMagnifyingGlass();
            TestTransparent();
        }

        private void TestRemainderEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            AssertIsSame(startIdx, startIdx + 2);
            AssertIsSame(startIdx + 1, startIdx + 3);
        }

        private void TestFrostedGlass()
        {
            PpOperations.SelectSlide(31);
            PpOperations.SelectShape("selectMe");
            PplFeatures.FrostedGlassEffect();
            AssertIsSame(31, 32);

            PpOperations.SelectSlide(29);
            PpOperations.SelectShape("selectMe");
            PplFeatures.FrostedGlassEffect();
            AssertIsSame(29, 30);
        }

        private void TestMagnifyingGlass()
        {
            PpOperations.SelectSlide(6);
            PpOperations.SelectShape("selectMe");
            PplFeatures.MagnifyingGlassEffect();
            AssertIsSame(6, 7);
        }

        private void TestTransparent()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShape("selectMe");
            PplFeatures.TransparentEffect();
            AssertIsSame(4, 5);
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
