using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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
            // check if timer is already created
            if (IsTimerExist())
            {
                ShowErrorMessageBox(TimerLabConstants.ErrorMessageOneTimerOnly);
                return;
            }

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
            timerBody.Name = TimerLabConstants.TimerBody;
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
            sliderHead.Name = TimerLabConstants.TimerSliderHead;
            sliderHead.Rotation = TimerLabConstants.Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = TimerLabConstants.TransparencyTranparent;

            var sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                timerLeft - (sliderBodyWidth / 2), timerTop, sliderBodyWidth, timerHeight);
            sliderBody.Name = TimerLabConstants.TimerSliderBody;
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

            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();
            var currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                // Add time marker
                var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight, 
                    timeMarkerWidth, timeMarkerHeight);
                timeMarker.Name = TimerLabConstants.TimerTimeMarker;
                timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
                timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;
                timeMarker.TextFrame.TextRange.Text = currentMarker.ToString();
                timeMarkers.Add(timeMarker);

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    var lineMarker = this.GetCurrentSlide().Shapes.AddShape
                        (Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                        timerLeft + (currentMarker * widthPerSec), timerTop, lineMarkerWidth, timerHeight);
                    lineMarker.Name = TimerLabConstants.TimerLineMarker;
                    lineMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                    lineMarkers.Add(timeMarker);
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

            if (lineMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerLineMarker, TimerLabConstants.TimerLineMarkerGroup);
            }
            if (timeMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerTimeMarker, TimerLabConstants.TimerTimeMarkerGroup);
            }
        }

        private void AddMinutesMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();
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
                    timeMarker.Name = TimerLabConstants.TimerTimeMarker;
                    timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
                    timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                    timeMarker.TextFrame.TextRange.Font.Color.RGB = 0;
                    timeMarkers.Add(timeMarker);

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
                    lineMarker.Name = TimerLabConstants.TimerLineMarker;
                    lineMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
                    lineMarkers.Add(lineMarker);

                    // set line width back to normal
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

            if (lineMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerLineMarker, TimerLabConstants.TimerLineMarkerGroup);
            }
            if (timeMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerTimeMarker, TimerLabConstants.TimerTimeMarkerGroup);
            }
        }

        private void AddSliderMotionEffect(Shape shape, int duration, float timerWidth, float slideWidth, 
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

        private void AddSliderEndEffect(Shape shape, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Effect sliderEndEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectDarken, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            sliderEndEffect.Timing.Duration = TimerLabConstants.ColorChangeDuration;
        }

        # region NumericUpDown Control
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

            int duration = Duration();
            Shape lineMarkerGroup = GetShapeByName(TimerLabConstants.TimerLineMarkerGroup);
            lineMarkerGroup.Delete();
            Shape timerMarkeGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroup);
            timerMarkeGroup.Delete();
            if (duration <= TimerLabConstants.SecondsInMinute)
            {
                Shape timerBody = GetShapeByName(TimerLabConstants.TimerBody);
                AddSecondsMarker(duration, TimerLabConstants.DefaultDenomination,
                    timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top,
                    TimerLabConstants.DefaultSecondsLineMarkerWidth, TimerLabConstants.DefaultTimeMarkerWidth,
                    TimerLabConstants.DefaultTimeMarkerHeight);
            }
            else
            {
                Shape timerBody = GetShapeByName(TimerLabConstants.TimerBody);
                AddMinutesMarker(duration, TimerLabConstants.DefaultDenomination,
                    timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top,
                    TimerLabConstants.DefaultSecondsLineMarkerWidth, TimerLabConstants.DefaultTimeMarkerWidth,
                    TimerLabConstants.DefaultTimeMarkerHeight);
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
            // update text box value
            float value = (float)WidthSlider.Value;
            WidthTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            var timerBody = GetShapeByName(TimerLabConstants.TimerBody);
            if (timerBody != null)
            {
                float increment = value - timerBody.Width;

                var lineMarkerGroup = GetShapeByName(TimerLabConstants.TimerLineMarkerGroup);
                var timeMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroup);
                timerBody.Select();
                lineMarkerGroup.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
                timeMarkerGroup.Select(Microsoft.Office.Core.MsoTriState.msoFalse);

                var tempGroup = this.GetCurrentSelection().ShapeRange.Group();
                tempGroup.Left = NewPosition(tempGroup.Left, increment);
                tempGroup.Width = tempGroup.Width + increment;

                var ungroupedShapes = tempGroup.Ungroup();
                timerBody = GetShapeByName(TimerLabConstants.TimerBody, ungroupedShapes);
                var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHead);
                var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBody);
                sliderBody.Left = NewPosition(timerBody.Left, sliderBody.Width);
                sliderHead.Left = NewPosition(timerBody.Left, sliderHead.Width);

                // update animation
                foreach (PowerPoint.Effect effect in this.GetCurrentSlide().TimeLine.MainSequence)
                {
                    if (effect.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectPathRight)
                    {
                        if (effect.Shape.Name.Equals(TimerLabConstants.TimerSliderBody) || 
                            effect.Shape.Name.Equals(TimerLabConstants.TimerSliderHead))
                        {
                            float end = timerBody.Width / this.GetCurrentPresentation().SlideWidth;
                            effect.Behaviors[1].MotionEffect.Path = "M 0 0 L " + end + " 0 E";
                        }
                    }
                }
            }
        }

        private void WidthTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            WidthTextBox.Text = TimerLabConstants.DefaultTimerWidth.ToString();
        }

        private void WidthTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = IsNumbersOnly(e.Text);
        }

        private void WidthTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(WidthTextBox.Text))
            {
                WidthTextBox.Text = ((int)WidthSlider.Value).ToString();
                return;
            }

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
            // update text box value
            float value = (float)HeightSlider.Value;
            HeightTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            var timerBody = GetShapeByName(TimerLabConstants.TimerBody);
            if (timerBody != null)
            {
                float increment = value - timerBody.Height;
                timerBody.Top = NewPosition(timerBody.Top, increment);
                timerBody.Height = timerBody.Height + increment;

                var lineMarkerGroup = GetShapeByName(TimerLabConstants.TimerLineMarkerGroup);
                lineMarkerGroup.Top = NewPosition(lineMarkerGroup.Top, increment);
                lineMarkerGroup.Height = lineMarkerGroup.Height + increment;

                var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBody);
                sliderBody.Top = NewPosition(sliderBody.Top, increment);
                sliderBody.Height = sliderBody.Height + increment;

                var timeMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroup);
                timeMarkerGroup.Top = NewPosition(timeMarkerGroup.Top, -increment);

                var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHead);
                sliderHead.Top = NewPosition(sliderHead.Top, increment);
            }
        }

        private void HeightTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            HeightTextBox.Text = TimerLabConstants.DefaultTimerHeight.ToString();
        }

        private void HeightTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = IsNumbersOnly(e.Text);
        }

        private void HeightTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(HeightTextBox.Text))
            {
                HeightTextBox.Text = ((int)HeightSlider.Value).ToString();
                return;
            }

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

        #region Shapes Helper
        private float NewPosition(float originalPosition, float objectSize)
        {
            return originalPosition - objectSize / 2;
        }

        private Shape GetShapeByName(string name)
        {
            return GetShapeByName(name, this.GetCurrentSlide().Shapes);
        }

        private Shape GetShapeByName(string name, dynamic shapes)
        {
            foreach (Shape shape in shapes)
            {
                if (shape.Name.Equals(name))
                {
                    return shape;
                }
            }
            return null;
        }

        private void GroupShapes(string shapeName, string groupName)
        {
            bool firstInGroup = true;
            foreach (Shape shape in this.GetCurrentSlide().Shapes)
            {
                if (shape.Name.Equals(shapeName))
                {
                    if (firstInGroup)
                    {
                        shape.Select();
                        firstInGroup = false;
                    }
                    else
                    {
                        shape.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
                    }
                }
            }
            Shape group = this.GetCurrentSelection().ShapeRange.Group();
            group.Name = groupName;
        }
        #endregion

        #region Validation Helper
        private bool IsTimerExist()
        {
            var timerBody = GetShapeByName(TimerLabConstants.TimerBody);
            return timerBody != null;
        }

        private bool IsNumbersOnly(string text)
        {
            Regex regex = new Regex("[^0-9]+");
            return regex.IsMatch(text);
        }
        #endregion

        #region Error Handling
        private void ShowErrorMessageBox(string content)
        {
            MessageBox.Show(content, "Error");
        }
        #endregion
    }
}
