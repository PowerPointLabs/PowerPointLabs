using System;
using System.Drawing;
using System.Windows.Forms;
using TestInterface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using Point = System.Drawing.Point;

namespace Test.FunctionalTest
{
    [TestClass]
    public class TimerLabTest : BaseFunctionalTest
    {
        private const int OriginalSlideNo = 4;
        private const int DefaultTimerSlideNo = 5;
        private const int ChangeWidthSlideNo = 6;
        private const int ChangeHeightSlideNo = 7;
        private const int RecreateBodySlideNo = 8;
        private const int ChangeLineColorSlideNo = 9;
        private const int ChangeDurationSlideNo = 10;
        private const int RecreateTimerBarSlideNo = 11;
        private const int CheckLineColorAndTextColorSlideNo = 12;
        private const int DurationInvalidSlideNo = 13;

        protected override string GetTestingSlideName()
        {
            return "TimerLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_TimerLabTest()
        {
            var timerLab = PplFeatures.TimerLab;
            timerLab.OpenPane();

            TestCreateDefaultTimer(timerLab);
            TestEditTimerWidth(timerLab);
            TestEditTimerHeight(timerLab);
            // TODO: the rest of the tests (After bug is fixed)
        }

        private void TestCreateDefaultTimer(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.ClickCreateButton();
            AssertIsSame(OriginalSlideNo, DefaultTimerSlideNo);
        }

        private void TestEditTimerWidth(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetWidthTextBoxValue(200);
            AssertIsSame(OriginalSlideNo, ChangeWidthSlideNo);
        }

        private void TestEditTimerHeight(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetHeightTextBoxValue(450);
            AssertIsSame(OriginalSlideNo, ChangeHeightSlideNo);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideNo);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
            //Need to check exact items color and etc.
        }
    }
}
