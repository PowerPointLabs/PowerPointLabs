using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AutoCaptionsTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "CaptionsLab\\AutoCaptions.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AutoCaptionsTest()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(4);
            ThreadUtil.WaitFor(1000);

            PplFeatures.AutoCaptions();

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(5);
            PpOperations.SelectShape("text 3").SafeDelete();

            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CaptionsMessageOneEmptySlide()
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(6);
            ThreadUtil.WaitFor(1000);

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                CaptionsLabText.ErrorDialogTitle,
                "Captions could not be created because there are no notes entered. Please enter something in the notes and try again.",
                PplFeatures.AutoCaptions);
        }
    }
}
