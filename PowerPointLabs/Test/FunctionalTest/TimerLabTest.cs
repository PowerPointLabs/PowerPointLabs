using System;
using System.Drawing;
using System.Windows.Forms;
using TestInterface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using Point = System.Drawing.Point;
using System.Collections.Generic;

namespace Test.FunctionalTest
{
    [TestClass]
    public class TimerLabTest : BaseFunctionalTest
    {
        private const int OriginalSlideNo = 4;
        private const int InitialTimerSlideNo = 5;
        private const int ChangeWidthSlideNo = 6;
        private const int ChangeHeightSlideNo = 7;
        private const int RecreateBodySlideNo = 8;
        private const int ChangeLineColorAndRecreateTimeMarkerSlideNo = 9;
        private const int ChangeDurationSlideNo = 10;
        private const int ChangeTextColorAndRecreateSliderSlideNo = 11;
        private const int ChangeDurationAndWidthSlideNo = 12;
        private const int DurationInvalidSlideNo = 13;

        private const string TimerBody = "TimerBody";
        private const string TimerLineMarkerGroup = "TimerLineMarkerGroup";
        private const string TimerTimeMarkerGroup = "TimerTimeMarkerGroup";
        private const string TimerSliderBody = "TimerSliderBody";
        private const string TimerSliderHead = "TimerSliderHead";


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

            TestCreateInitialTimer(timerLab);
            TestEditTimerWidth(timerLab);
            TestEditTimerHeight(timerLab);
            TestDeleteTimerBody(timerLab);
            TestEditLineColorAndDeleteTimeMarker(timerLab);
            TestEditTimerDuration(timerLab);
            TestEditTextColorAndDeleteSlider(timerLab);
            TestEditDurationAndWidth(timerLab);
            TestInvalidDuration(timerLab);
        }

        private void TestCreateInitialTimer(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(1.05);
            timerLab.ClickCreateButton();
            AssertIsSame(OriginalSlideNo, InitialTimerSlideNo);
        }

        private void TestEditTimerWidth(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetWidthTextBoxValue(250);
            AssertIsSame(OriginalSlideNo, ChangeWidthSlideNo);
        }

        private void TestEditTimerHeight(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetHeightTextBoxValue(400);
            AssertIsSame(OriginalSlideNo, ChangeHeightSlideNo);
        }

        private void TestDeleteTimerBody(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            var shapes = PpOperations.SelectShape(TimerBody);
            shapes.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
               "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(OriginalSlideNo, RecreateBodySlideNo);
        }

        private void TestEditLineColorAndDeleteTimeMarker(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(ChangeLineColorAndRecreateTimeMarkerSlideNo);
            int expectedColor = PpOperations.SelectShape(TimerLineMarkerGroup)[1].Line.ForeColor.RGB;

            PpOperations.SelectSlide(OriginalSlideNo);
            var lineMarkerGroup = PpOperations.SelectShape(TimerLineMarkerGroup);
            lineMarkerGroup.Line.ForeColor.RGB = expectedColor;
            var timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
              "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(OriginalSlideNo, ChangeLineColorAndRecreateTimeMarkerSlideNo);
        }

        private void TestEditTimerDuration(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(0.07);
            AssertIsSame(OriginalSlideNo, ChangeDurationSlideNo);
        }

        private void TestEditTextColorAndDeleteSlider(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(ChangeTextColorAndRecreateSliderSlideNo);
            int expectedColor = PpOperations.SelectShape(TimerTimeMarkerGroup)[1].TextFrame.TextRange.Font.Color.RGB;

            PpOperations.SelectSlide(OriginalSlideNo);
            var timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.TextFrame.TextRange.Font.Color.RGB = expectedColor;
            List<string> sliderComponentNames = new List<string> { TimerSliderHead, TimerSliderBody };
            var sliderComponents = PpOperations.SelectShapes(sliderComponentNames);
            sliderComponents.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
              "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(OriginalSlideNo, ChangeTextColorAndRecreateSliderSlideNo);
        }

        private void TestEditDurationAndWidth(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(4.56);
            timerLab.SetWidthSliderValue(654);

            AssertIsSame(OriginalSlideNo, ChangeDurationAndWidthSlideNo);
        }

        private void TestInvalidDuration(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(5.67);

            AssertIsSame(OriginalSlideNo, DurationInvalidSlideNo);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideNo);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

            SlideUtil.IsSameShapes(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
