using System;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TimerLab
{
    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class TimerLabPaneWPF : UserControl
    {
        public const int DefaultDenomination = 10;
        public const float DefaultTimerWidth = 500;
        public const float DefaultTimerHeight = 50;
        public const float DefaultMinutesLineMarkerWidth = 4;
        public const float DefaultSecondsLineMarkerWidth = 2;
        public const float DefaultTimeMarkerHeight = 50;
        public const float DefaultTimeMarkerWidth = 50;
        public const float DefaultSliderBodyWidth = 4;
        public const float DefaultSliderHeadSize = 20;

        public const int TransparencyTranparent = 1;
        public const int Rotate180Degrees = 180;
        public const float ColorChangeDuration = 0.001f;

        public const int SecondsInMinute = 60;
        public const int StartTime = 0;
        public const double FractionalDecrementOffset = 0.60;
        public const double FractionalDecrementLowerBound = 0.00;
        public const double FractionalIncrementOffset = 0.99;
        public const double FractionalIncrementUpperBound = 0.59;

        public TimerLabPaneWPF()
        {
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // Slide dimensions
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;

            // Functionality
            int duration = Duration();
            int denomination = DefaultDenomination;

            // Aesthetics
            // Main
            float timerWidth = TimerWidth();
            float timerHeight = TimerHeight();
            int timerBodyColor = System.Drawing.Color.FromArgb(106, 84, 68).ToArgb();
            // Markers
            float lineMarkerWidth = DefaultSecondsLineMarkerWidth;
            float timeMarkerWidth = DefaultTimeMarkerWidth;
            float timeMarkerHeight = DefaultTimeMarkerHeight;
            // Slider
            var sliderHeadSize = DefaultSliderHeadSize;
            var sliderBodyWidth = DefaultSliderBodyWidth;
            int sliderColor = System.Drawing.Color.FromArgb(70, 150, 247).ToArgb();

            // Position
            float timerLeft = (slideWidth - timerWidth) / 2;
            float timerTop = (slideHeight - timerHeight) / 2;         

            // Create timer
            var timerBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                timerLeft, timerTop, timerWidth, timerHeight);
            timerBody.Name = "timerBody";
            timerBody.Fill.ForeColor.RGB = timerBodyColor;
            timerBody.Line.ForeColor.RGB = timerBodyColor;
            
            // Add markers
            if (duration <= SecondsInMinute)
            {
                AddSecondsMarker(duration, denomination, timerWidth, timerHeight, timerLeft, timerTop,
                    lineMarkerWidth, timeMarkerWidth, timeMarkerHeight);
            }
            else
            {
                AddMinutesMarker(duration, denomination, timerWidth, timerHeight, timerLeft, timerTop,
                    lineMarkerWidth, timeMarkerWidth, timeMarkerHeight);
            }

            // Add slider components
            var sliderHead = this.GetCurrentSlide().Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, 
                timerLeft - (sliderHeadSize / 2), timerTop - (sliderHeadSize / 2), sliderHeadSize, sliderHeadSize);
            sliderHead.Name = "timerSliderHead";
            sliderHead.Rotation = Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = TransparencyTranparent;

            var sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                timerLeft - (sliderBodyWidth / 2), timerTop, sliderBodyWidth, timerHeight);
            sliderBody.Name = "timerSliderBody";
            sliderBody.Fill.ForeColor.RGB = sliderColor;
            sliderBody.Line.Transparency = TransparencyTranparent;

            // Add slider animations
            AddSliderMotionEffect(sliderHead, duration, timerWidth, slideWidth, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            AddSliderMotionEffect(sliderBody, duration, timerWidth, slideWidth, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            AddSliderEndEffect(sliderHead, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            AddSliderEndEffect(sliderBody, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
        }

        private int Duration()
        {
            int duration = SecondsInMinute;
            if (DurationTextBox.Value != null)
            {
                double value = Math.Round(DurationTextBox.Value.Value, 2);
                int minutes = (int)value;
                int seconds = (int)((value - minutes) * 100);
                duration = (minutes * SecondsInMinute) + seconds;
            }
            return duration;
        }

        private float TimerWidth()
        {
            float width = DefaultTimerWidth;
            if (!string.IsNullOrEmpty(WidthTextBox.Text))
            {
                width = float.Parse(WidthTextBox.Text);
            }
            return width;
        }

        private float TimerHeight()
        {
            float height = DefaultTimerHeight;
            if (!string.IsNullOrEmpty(HeightTextBox.Text))
            {
                height = float.Parse(HeightTextBox.Text);
            }
            return height;
        }

        private void AddSecondsMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            var currentMarker = StartTime;
            while (currentMarker <= duration)
            {
                // Add time marker
                var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight, 
                    timeMarkerWidth, timeMarkerHeight);
                timeMarker.Name = "timerTimeMarker";
                timeMarker.Fill.Transparency = TransparencyTranparent;
                timeMarker.Line.Transparency = TransparencyTranparent;
                timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;
                timeMarker.TextFrame.TextRange.Text = currentMarker.ToString();

                // Add line marker if it is not the start or end
                if (currentMarker != StartTime && currentMarker != duration)
                {
                    var lineMarker = this.GetCurrentSlide().Shapes.AddShape
                        (Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        timerLeft + (currentMarker * widthPerSec), timerTop, lineMarkerWidth, timerHeight);
                    lineMarker.Name = "timerLineMarker";
                    lineMarker.Line.Transparency = TransparencyTranparent;
                }

                if (currentMarker == duration)
                {
                    break;
                }

                currentMarker += denomination;
                if (currentMarker > duration)
                {
                    currentMarker = duration;
                }
            }
        }

        private void AddMinutesMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            var currentMarker = StartTime;
            while (currentMarker <= duration)
            {
                // Add time markers for start, end and every minute
                if (currentMarker % SecondsInMinute == 0 || currentMarker == duration)
                {
                    var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight, 
                        timeMarkerWidth, timeMarkerHeight);
                    timeMarker.Name = "timerTimeMarker";
                    timeMarker.Fill.Transparency = TransparencyTranparent;
                    timeMarker.Line.Transparency = TransparencyTranparent;
                    timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;

                    int remainingSeconds = currentMarker % SecondsInMinute;
                    if (currentMarker == duration && remainingSeconds != 0)
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / SecondsInMinute).ToString() + ":" + remainingSeconds.ToString();
                    }
                    else
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / SecondsInMinute).ToString();
                    }
                }

                // Add line marker if it is not the start or end
                if (currentMarker != StartTime && currentMarker != duration)
                {
                    //Thicken the line if it is a minute marker
                    if (currentMarker % SecondsInMinute == 0)
                    {
                        lineMarkerWidth = DefaultMinutesLineMarkerWidth;
                    }
                    var lineMarker = this.GetCurrentSlide().Shapes.AddShape(
                        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        timerLeft + (currentMarker * widthPerSec), timerTop, lineMarkerWidth, timerHeight);
                    lineMarker.Name = "timerLineMarker";
                    lineMarker.Line.Transparency = TransparencyTranparent;
                    lineMarkerWidth = DefaultSecondsLineMarkerWidth;
                }

                if (currentMarker == duration)
                {
                    break;
                }

                currentMarker += denomination;
                if (currentMarker > duration)
                {
                    currentMarker = duration;
                }
            }
        }

        private void AddSliderMotionEffect(PowerPoint.Shape shape, int duration, float timerWidth, float slideWidth, 
            PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Effect sliderMotionEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectPathRight, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
            PowerPoint.AnimationBehavior motion = sliderMotionEffect.Behaviors[1];
            float end = timerWidth / slideWidth;
            motion.MotionEffect.Path = "M 0 0 L " + end + " 0 E";
            sliderMotionEffect.Timing.Duration = duration;
            sliderMotionEffect.Timing.SmoothStart = Microsoft.Office.Core.MsoTriState.msoFalse;
            sliderMotionEffect.Timing.SmoothEnd = Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        private void AddSliderEndEffect(PowerPoint.Shape shape, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Effect sliderEndEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectDarken, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            sliderEndEffect.Timing.Duration = ColorChangeDuration;
        }

        # region NumericUpDown Control Customisation
        private void DurationTextBox_ValueDecremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                return;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == FractionalDecrementLowerBound)
            {
                DurationTextBox.Value = (integerPart - 1) + FractionalDecrementOffset;
            }
        }

        private void DurationTextBox_ValueIncremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                return;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + FractionalIncrementOffset;
            }
        }

        private void DurationTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (DurationTextBox.Value == null)
            {
                return;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) > FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + 1;
            }
        }

        #endregion
    }
}
