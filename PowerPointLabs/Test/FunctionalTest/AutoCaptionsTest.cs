using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoCaptionsTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AutoCaptions.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
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

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CaptionsMessageOneEmptySlide()
        {
            var actualSlide = PpOperations.SelectSlide(6);
            ThreadUtil.WaitFor(1000);

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                TextCollection.CaptionsLabErrorDialogTitle,
                "Captions could not be created because there are no notes entered. Please enter something in the notes and try again.",
                PplFeatures.AutoCaptions);
        }
    }
}
