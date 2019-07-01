using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Util;

using TestInterface;

namespace Test.FunctionalTest
{
    [TestClass]
    public class TimerLabTest : BaseFunctionalTest
    {
        // Original Timer Lab
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
        private const int CountdownCheckedSlideNo = 14;
        private const int ChangeDurationWithCountdownSlideNo = 15;
        private const int CountdownAndNonMultipleDenominationDurationSlideNo = 16;
        private const int ProgressBarCheckedSlideNo = 17;

        // Timer Lab Progress Bar
        private const int PbOriginalSlideNo = 20;
        private const int PbInitialTimerSlideNo = 21;
        private const int PbChangeWidthSlideNo = 22;
        private const int PbChangeHeightSlideNo = 23;
        private const int PbRecreateBodySlideNo = 24;
        private const int PbChangeLineColorAndRecreateTimeMarkerSlideNo = 25;
        private const int PbChangeDurationSlideNo = 26;
        private const int PbChangeTextColorAndRecreateSliderSlideNo = 27;
        private const int PbChangeDurationAndWidthSlideNo = 28;
        private const int PbDurationInvalidSlideNo = 29;
        private const int PbCountdownCheckedSlideNo = 30;
        private const int PbChangeDurationWithCountdownSlideNo = 31;
        private const int PbCountdownAndNonMultipleDenominationDurationSlideNo = 32;
        private const int PbProgressBarUncheckedSlideNo = 33;

        private const string TimerBody = "TimerBody";
        private const string TimerLineMarkerGroup = "TimerLineMarkerGroup";
        private const string TimerTimeMarkerGroup = "TimerTimeMarkerGroup";
        private const string TimerSliderBody = "TimerSliderBody";
        private const string TimerSliderHead = "TimerSliderHead";
        private const string ProgressBar = "ProgressBar";


        protected override string GetTestingSlideName()
        {
            return "TimerLab\\TimerLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_TimerLabTest()
        {
            ITimerLabController timerLab = PplFeatures.TimerLab;
            timerLab.OpenPane();

            // Original Timer Lab
            TestCreateInitialTimer(timerLab);
            TestEditTimerWidth(timerLab);
            TestEditTimerHeight(timerLab);
            TestDeleteTimerBody(timerLab);
            TestEditLineColorAndDeleteTimeMarker(timerLab);
            TestEditTimerDuration(timerLab);
            TestEditTextColorAndDeleteSlider(timerLab);
            TestEditDurationAndWidth(timerLab);
            TestInvalidDuration(timerLab);
            TestEditCountdownState(timerLab);
            TestEditDurationWithCountdownTimer(timerLab);
            TestNonMultipleDenominationDurationWithCountdownTimer(timerLab);
            TestEditProgressBarState(timerLab);

            RevertSettingsToOriginal(timerLab);
            // Timer Lab Progress Bar
            TestCreateInitialTimerPb(timerLab);
            TestEditTimerWidthPb(timerLab);
            TestEditTimerHeightPb(timerLab);
            TestDeleteTimerBodyPb(timerLab);
            TestEditLineColorAndDeleteTimeMarkerPb(timerLab);
            TestEditTimerDurationPb(timerLab);
            TestEditTextColorAndDeleteProgressBarPb(timerLab);
            TestEditDurationAndWidthPb(timerLab);
            TestInvalidDurationPb(timerLab);
            TestEditCountdownStatePb(timerLab);
            TestEditDurationWithCountdownTimerPb(timerLab);
            TestNonMultipleDenominationDurationWithCountdownTimerPb(timerLab);
            TestEditProgressBarStatePb(timerLab);
        }

        private void TestCreateInitialTimer(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(1.05);
            timerLab.SetProgressBarCheckBoxState(false);
            timerLab.SetCountdownCheckBoxState(false);
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
            ShapeRange shapes = PpOperations.SelectShape(TimerBody);
            shapes.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
               "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(OriginalSlideNo, RecreateBodySlideNo);
        }

        private void TestEditLineColorAndDeleteTimeMarker(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(ChangeLineColorAndRecreateTimeMarkerSlideNo);
            int expectedColor = PpOperations.SelectShape(TimerLineMarkerGroup)[1].Line.ForeColor.RGB;

            PpOperations.SelectSlide(OriginalSlideNo);
            ShapeRange lineMarkerGroup = PpOperations.SelectShape(TimerLineMarkerGroup);
            lineMarkerGroup.Line.ForeColor.RGB = expectedColor;
            ShapeRange timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.SafeDelete();

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
            ShapeRange timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.TextFrame.TextRange.Font.Color.RGB = expectedColor;
            List<string> sliderComponentNames = new List<string> { TimerSliderHead, TimerSliderBody };
            ShapeRange sliderComponents = PpOperations.SelectShapes(sliderComponentNames);
            sliderComponents.SafeDelete();

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

        private void TestEditCountdownState(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetCountdownCheckBoxState(true);

            AssertIsSame(OriginalSlideNo, CountdownCheckedSlideNo);
        }

        private void TestEditDurationWithCountdownTimer(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(0.30);

            AssertIsSame(OriginalSlideNo, ChangeDurationWithCountdownSlideNo);
        }

        private void TestNonMultipleDenominationDurationWithCountdownTimer(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetDurationTextBoxValue(4.16);

            AssertIsSame(OriginalSlideNo, CountdownAndNonMultipleDenominationDurationSlideNo);
        }

        private void TestEditProgressBarState(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(OriginalSlideNo);
            timerLab.SetProgressBarCheckBoxState(true);

            AssertIsSame(OriginalSlideNo, ProgressBarCheckedSlideNo);
        }

        private void TestCreateInitialTimerPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetProgressBarCheckBoxState(true);
            timerLab.SetDurationTextBoxValue(1.05);
            timerLab.ClickCreateButton();
            AssertIsSame(PbOriginalSlideNo, PbInitialTimerSlideNo);
        }

        private void TestEditTimerWidthPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetWidthTextBoxValue(250);
            AssertIsSame(PbOriginalSlideNo, PbChangeWidthSlideNo);
        }

        private void TestEditTimerHeightPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetHeightTextBoxValue(400);
            AssertIsSame(PbOriginalSlideNo, PbChangeHeightSlideNo);
        }

        private void TestDeleteTimerBodyPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            ShapeRange shapes = PpOperations.SelectShape(TimerBody);
            shapes.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
               "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(PbOriginalSlideNo, PbRecreateBodySlideNo);
        }

        private void TestEditLineColorAndDeleteTimeMarkerPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(ChangeLineColorAndRecreateTimeMarkerSlideNo);
            int expectedColor = PpOperations.SelectShape(TimerLineMarkerGroup)[1].Line.ForeColor.RGB;

            PpOperations.SelectSlide(PbOriginalSlideNo);
            ShapeRange lineMarkerGroup = PpOperations.SelectShape(TimerLineMarkerGroup);
            lineMarkerGroup.Line.ForeColor.RGB = expectedColor;
            ShapeRange timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
              "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(PbOriginalSlideNo, PbChangeLineColorAndRecreateTimeMarkerSlideNo);
        }

        private void TestEditTimerDurationPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetDurationTextBoxValue(0.07);
            AssertIsSame(PbOriginalSlideNo, PbChangeDurationSlideNo);
        }

        private void TestEditTextColorAndDeleteProgressBarPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(ChangeTextColorAndRecreateSliderSlideNo);
            int expectedColor = PpOperations.SelectShape(TimerTimeMarkerGroup)[1].TextFrame.TextRange.Font.Color.RGB;

            PpOperations.SelectSlide(PbOriginalSlideNo);
            ShapeRange timeMarkerGroup = PpOperations.SelectShape(TimerTimeMarkerGroup);
            timeMarkerGroup.TextFrame.TextRange.Font.Color.RGB = expectedColor;
            ShapeRange progressBar = PpOperations.SelectShape(ProgressBar);
            progressBar.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp("Error",
              "Only one timer allowed per slide.", timerLab.ClickCreateButton, "Ok");
            AssertIsSame(PbOriginalSlideNo, PbChangeTextColorAndRecreateSliderSlideNo);
        }

        private void TestEditDurationAndWidthPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetDurationTextBoxValue(4.56);
            timerLab.SetWidthSliderValue(654);

            AssertIsSame(PbOriginalSlideNo, PbChangeDurationAndWidthSlideNo);
        }

        private void TestInvalidDurationPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetDurationTextBoxValue(5.67);

            AssertIsSame(PbOriginalSlideNo, PbDurationInvalidSlideNo);
        }

        private void TestEditCountdownStatePb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetCountdownCheckBoxState(true);

            AssertIsSame(PbOriginalSlideNo, PbCountdownCheckedSlideNo);
        }

        private void TestEditDurationWithCountdownTimerPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetDurationTextBoxValue(0.30);

            AssertIsSame(PbOriginalSlideNo, PbChangeDurationWithCountdownSlideNo);
        }

        private void TestNonMultipleDenominationDurationWithCountdownTimerPb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetDurationTextBoxValue(4.16);

            AssertIsSame(PbOriginalSlideNo, PbCountdownAndNonMultipleDenominationDurationSlideNo);
        }

        private void TestEditProgressBarStatePb(ITimerLabController timerLab)
        {
            PpOperations.SelectSlide(PbOriginalSlideNo);
            timerLab.SetProgressBarCheckBoxState(false);

            AssertIsSame(PbOriginalSlideNo, PbProgressBarUncheckedSlideNo);
        }

        private void RevertSettingsToOriginal(ITimerLabController timerLab)
        {
            timerLab.SetDurationTextBoxValue(1.00);
            timerLab.SetCountdownCheckBoxState(false);
            timerLab.SetProgressBarCheckBoxState(false);
            timerLab.SetHeightSliderValue(50);
            timerLab.SetWidthSliderValue(600);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideNo);
            Slide expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

            SlideUtil.IsSameShapes(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }

    }
}
