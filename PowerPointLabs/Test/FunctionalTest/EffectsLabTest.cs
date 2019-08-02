using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class EffectsLabTest : BaseFunctionalTest
    {
        private const int ZIndexRecolorTestSlide = 54;
        private const int ZIndexBlurTestSlide = 51;
        private const int BlurSelectedShapeTestSlide = 49;
        private const int BlurSelectedShapeWithTextTestSlide = 47;
        private const int BlurSelectedTextBoxTestSide = 45;
        private const int BlurSelectedGroupTestSlide = 43;
        private const int BlurRemainderShapeTestSlide = 40;
        private const int BlurRemainderShapeWithTextTestSlide = 37;
        private const int BlurBackgroundTextBoxTestSlide = 34;
        private const int BlurBackgroundGroupTestSlide = 31;
        private const int RecolorSepiaTestSlide = 28;
        private const int RecolorGothamTestSlide = 25;
        private const int RecolorBlackAndWhiteTestSlide = 22;
        private const int RecolorGrayScaleTestSlide = 19;
        private const int TintBlurSelectedShapeTestSlide = 17;
        private const int TintBlurRemainderTextBoxTestSlide = 14;
        private const int TintBlurRemainderGroupTestSlide = 11;
        private const int TintBlurBackgroundShapeWithTextTestSlide = 8;
        private const int MagnifyTestSlide = 6;
        private const int TransparentTestSlide = 4;

        protected override string GetTestingSlideName()
        {
            return "EffectsLab\\EffectsLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_EffectsLabTest()
        {
            TestGeneratedSlideEffect(ZIndexRecolorTestSlide, PplFeatures.SepiaBackgroundEffect);
            TestGeneratedSlideEffect(ZIndexBlurTestSlide, PplFeatures.BlurBackgroundEffect);

            TestSlideEffect(BlurSelectedShapeTestSlide, PplFeatures.BlurSelectedEffect);
            TestSlideEffect(BlurSelectedShapeWithTextTestSlide, PplFeatures.BlurSelectedEffect);
            TestSlideEffect(BlurSelectedTextBoxTestSide, PplFeatures.BlurSelectedEffect);
            TestSlideEffect(BlurSelectedGroupTestSlide, PplFeatures.BlurSelectedEffect);

            TestGeneratedSlideEffect(BlurRemainderShapeTestSlide, PplFeatures.BlurRemainderEffect);
            TestGeneratedSlideEffect(BlurRemainderShapeWithTextTestSlide, PplFeatures.BlurRemainderEffect);
            TestGeneratedSlideEffect(BlurBackgroundTextBoxTestSlide, PplFeatures.BlurBackgroundEffect);
            TestGeneratedSlideEffect(BlurBackgroundGroupTestSlide, PplFeatures.BlurBackgroundEffect);

            // recolor
            TestGeneratedSlideEffect(RecolorSepiaTestSlide, PplFeatures.SepiaRemainderEffect);
            TestGeneratedSlideEffect(RecolorGothamTestSlide, PplFeatures.GothamRemainderEffect);
            TestGeneratedSlideEffect(RecolorBlackAndWhiteTestSlide, PplFeatures.BlackAndWhiteBackgroundEffect);
            TestGeneratedSlideEffect(RecolorGrayScaleTestSlide, PplFeatures.GrayScaleBackgroundEffect);

            // tinted
            PplFeatures.SetTintForBlurSelected(true);
            TestSlideEffect(TintBlurSelectedShapeTestSlide, PplFeatures.BlurSelectedEffect);

            PplFeatures.SetTintForBlurRemainder(true);
            TestGeneratedSlideEffect(TintBlurRemainderTextBoxTestSlide, PplFeatures.BlurRemainderEffect);
            TestGeneratedSlideEffect(TintBlurRemainderGroupTestSlide, PplFeatures.BlurRemainderEffect);

            PplFeatures.SetTintForBlurBackground(true);
            TestGeneratedSlideEffect(TintBlurBackgroundShapeWithTextTestSlide, PplFeatures.BlurBackgroundEffect);

            // other effects
            TestSlideEffect(MagnifyTestSlide, PplFeatures.MagnifyingGlassEffect);
            TestSlideEffect(TransparentTestSlide, PplFeatures.TransparentEffect);
        }

        private void TestGeneratedSlideEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            ThreadUtil.WaitFor(100);
            AssertIsSame(startIdx, startIdx + 2);
            AssertIsSame(startIdx + 1, startIdx + 3);
        }

        private void TestSlideEffect(int startIdx, Action effectAction)
        {
            PpOperations.SelectSlide(startIdx);
            PpOperations.SelectShape("selectMe");
            effectAction.Invoke();
            AssertIsSame(startIdx, startIdx + 1);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide, 0.99);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
