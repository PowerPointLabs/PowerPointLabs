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
            int denomination = TimerLabConstants.DefaultDenomination;

            // Aesthetics
            // Main
            float timerWidth = TimerWidth();
            float timerHeight = TimerHeight();
            int timerBodyColor = System.Drawing.Color.FromArgb(106, 84, 68).ToArgb();
            // Markers
            float lineMarkerWidth = TimerLabConstants.DefaultSecondsLineMarkerWidth;
            float timeMarkerWidth = TimerLabConstants.DefaultTimeMarkerWidth;
            float timeMarkerHeight = TimerLabConstants.DefaultTimeMarkerHeight;
            // Slider
            var sliderHeadSize = TimerLabConstants.DefaultSliderHeadSize;
            var sliderBodyWidth = TimerLabConstants.DefaultSliderBodyWidth;
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
            if (duration <= TimerLabConstants.SecondsInMinute)
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
            sliderHead.Rotation = TimerLabConstants.Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = TimerLabConstants.TransparencyTranparent;

            var sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                timerLeft - (sliderBodyWidth / 2), timerTop, sliderBodyWidth, timerHeight);
            sliderBody.Name = "timerSliderBody";
            sliderBody.Fill.ForeColor.RGB = sliderColor;
            sliderBody.Line.Transparency = TimerLabConstants.TransparencyTranparent;

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
            int duration = TimerLabConstants.SecondsInMinute;
            if (DurationTextBox.Value != null)
            {
                double value = Math.Round(DurationTextBox.Value.Value, 2);
                int minutes = (int)value;
                int seconds = (int)((value - minutes) * 100);
                duration = (minutes * TimerLabConstants.SecondsInMinute) + seconds;
            }
            return duration;
        }

        private float TimerWidth()
        {
            float width = (float)WidthSlider.Value;
            return width;
        }

        private float TimerHeight()
        {
            float height = (float)HeightSlider.Value;
            return height;
        }

        private void AddSecondsMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            var currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                // Add time marker
                var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight, 
                    timeMarkerWidth, timeMarkerHeight);
                timeMarker.Name = "timerTimeMarker";
                timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
                timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;
                timeMarker.TextFrame.TextRange.Text = currentMarker.ToString();

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    var lineMarker = this.GetCurrentSlide().Shapes.AddShape
                        (Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        timerLeft + (currentMarker * widthPerSec), timerTop, lineMarkerWidth, timerHeight);
                    lineMarker.Name = "timerLineMarker";
                    lineMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
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

            var currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                // Add time markers for start, end and every minute
                if (currentMarker % TimerLabConstants.SecondsInMinute == 0 || currentMarker == duration)
                {
                    var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight, 
                        timeMarkerWidth, timeMarkerHeight);
                    timeMarker.Name = "timerTimeMarker";
                    timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
                    timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                    timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;

                    int remainingSeconds = currentMarker % TimerLabConstants.SecondsInMinute;
                    if (currentMarker == duration && remainingSeconds != 0)
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / TimerLabConstants.SecondsInMinute).ToString() + ":"
                            + remainingSeconds.ToString();
                    }
                    else
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / TimerLabConstants.SecondsInMinute).ToString();
                    }
                }

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    //Thicken the line if it is a minute marker
                    if (currentMarker % TimerLabConstants.SecondsInMinute == 0)
                    {
                        lineMarkerWidth = TimerLabConstants.DefaultMinutesLineMarkerWidth;
                    }
                    var lineMarker = this.GetCurrentSlide().Shapes.AddShape(
                        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        timerLeft + (currentMarker * widthPerSec), timerTop, lineMarkerWidth, timerHeight);
                    lineMarker.Name = "timerLineMarker";
                    lineMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                    lineMarkerWidth = TimerLabConstants.DefaultSecondsLineMarkerWidth;
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
            sliderEndEffect.Timing.Duration = TimerLabConstants.ColorChangeDuration;
        }

        # region NumericUpDown Control Customisation
        private void DurationTextBox_ValueDecremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                DurationTextBox.Value = TimerLabConstants.DefaultDisplayDuration;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == TimerLabConstants.FractionalDecrementLowerBound)
            {
                DurationTextBox.Value = (integerPart - 1) + TimerLabConstants.FractionalDecrementOffset;
            }
        }

        private void DurationTextBox_ValueIncremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                DurationTextBox.Value = TimerLabConstants.DefaultDisplayDuration;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == TimerLabConstants.FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + TimerLabConstants.FractionalIncrementOffset;
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

            if (Math.Round(fractionalPart, 2) > TimerLabConstants.FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + 1;
            }
        }

        #endregion

        #region Width Control
        private void WidthSlider_Loaded(object sender, RoutedEventArgs e)
        {
            WidthSlider.Value = TimerLabConstants.DefaultTimerWidth;
            WidthSlider.Minimum = TimerLabConstants.MinTimerWidth;
            WidthSlider.Maximum = this.GetCurrentPresentation().SlideWidth;
           
        } 

        private void WidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            float value = (float)WidthSlider.Value;
            WidthTextBox.Text = ((int)value).ToString();
        }

        private void WidthTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            WidthTextBox.Text = TimerLabConstants.DefaultTimerWidth.ToString();
        }

        private void WidthTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            int value = Convert.ToInt32(WidthTextBox.Text);
            if (value < TimerLabConstants.MinTimerWidth)
            {
                value = (int)TimerLabConstants.MinTimerWidth;
            }
            else if (value > this.GetCurrentPresentation().SlideWidth)
            {
                value = (int)this.GetCurrentPresentation().SlideWidth;
            }
            WidthTextBox.Text = value.ToString();
            WidthSlider.Value = value;
        }
        #endregion

        #region Height Control
        private void HeightSlider_Loaded(object sender, RoutedEventArgs e)
        {
            HeightSlider.Minimum = TimerLabConstants.MinTimerHeight;
            HeightSlider.Maximum = this.GetCurrentPresentation().SlideHeight;
            HeightSlider.Value = TimerLabConstants.DefaultTimerHeight;
        }

        private void HeightSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            float value = (float)HeightSlider.Value;
            HeightTextBox.Text = ((int)value).ToString();        
        }

        private void HeightTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            HeightTextBox.Text = TimerLabConstants.DefaultTimerHeight.ToString();
        }

        private void HeightTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = HeightTextBox.Text;
            if (text.Length > TimerLabConstants.SizeStringLimit)
            {
                HeightTextBox.Text = text.Substring(0, text.Length - 1);
            }
        }

        private void HeightTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            int value = Convert.ToInt32(HeightTextBox.Text);
            if (value < TimerLabConstants.MinTimerHeight)
            {
                value = (int)TimerLabConstants.MinTimerHeight;
            }
            else if (value > this.GetCurrentPresentation().SlideHeight)
            {
                value = (int)this.GetCurrentPresentation().SlideHeight;
            }
            HeightTextBox.Text = value.ToString();
            HeightSlider.Value = value;
        }
        #endregion
    }
}
