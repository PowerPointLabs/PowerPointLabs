using System;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class EffectsLabTest : BaseFunctionalTest
    {
        [TestMethod]
        public void FT_EffectsLabTest()
        {
            TestSepia();
            TestGotham();
            TestBlackAndWhite();
            TestGreyScale();
            TestBlurRemainder();
            TestMagnifyingGlass();
            TestTransparent();
        }

        private void TestSepia()
        {
            TestRemainderEffect(20, PplFeatures.SepiaEffect);
        }

        private void TestGotham()
        {
            TestRemainderEffect(17, PplFeatures.GothamEffect);
        }

        private void TestBlackAndWhite()
        {
            TestRemainderEffect(14, PplFeatures.BlackAndWhiteEffect);
        }

        private void TestGreyScale()
        {
            TestRemainderEffect(11, PplFeatures.GreyScaleEffect);
        }

        private void TestBlurRemainder()
        {
            TestRemainderEffect(8, PplFeatures.BlurRemainderEffect);
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

        protected override string GetTestingSlideName()
        {
            return "EffectsLab.pptx";
        }

        private void TestRemainderEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            AssertIsSame(startIdx, startIdx + 2);
            AssertIsSame(startIdx + 1, startIdx + 3);
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
