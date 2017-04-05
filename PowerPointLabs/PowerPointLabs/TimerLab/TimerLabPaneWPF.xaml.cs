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
            if (HasTimer())
            {
                ReformMissingComponents();
                ShowErrorMessageBox(TimerLabConstants.ErrorMessageOneTimerOnly);
                return;
            }

            // Properties
            int duration = Duration();
            float timerWidth = TimerWidth();
            float timerHeight = TimerHeight();
            int timerBodyColor = TimerBodyColor();
            int sliderColor = SliderColor();

            // Position
            float timerLeft = DefaultTimerLeft(SlideWidth(), timerWidth);
            float timerTop = DefaultTimerTop(SlideHeight(), timerHeight);

            CreateBlocksTimer(duration, timerWidth, timerHeight, timerLeft, timerTop, timerBodyColor, sliderColor);
        }

        #region Slide Dimensions
        private float SlideWidth()
        {
            return this.GetCurrentPresentation().SlideWidth;
        }

        private float SlideHeight()
        {
            return this.GetCurrentPresentation().SlideHeight;
        }
        #endregion

        #region Timer Properties
        private int Duration()
        {
            int duration = TimerLabConstants.SecondsInMinute;
            if (DurationTextBox.Value != null)
            {
                double value = Math.Round(DurationTextBox.Value.Value, 2);
                int minutes = (int)value;
                int seconds = (int)(Math.Round(value - minutes, 2) * 100);
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

        private float DefaultTimerLeft(float slideWidth, float timerWidth)
        {
            return (slideWidth - timerWidth) / 2;
        }

        private float DefaultTimerTop(float slideHeight, float timerHeight)
        {
            return (slideHeight - timerHeight) / 2;
        }

        private int TimerBodyColor()
        {
            return System.Drawing.Color.FromArgb(106, 84, 68).ToArgb();
        }

        private int SliderColor()
        {
            return System.Drawing.Color.FromArgb(70, 150, 247).ToArgb();
        }

        private int TimeMarkerColor()
        {
            return System.Drawing.Color.FromArgb(90, 90, 90).ToArgb();
        }
        #endregion

        #region Timer Creation
        private void CreateBlocksTimer(int duration, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, int timerBodyColor, int sliderColor)
        {
            AddTimerBody(timerWidth, timerHeight, timerLeft, timerTop, timerBodyColor);
            AddMarkers(duration, timerWidth, timerHeight, timerLeft, timerTop);
            AddSlider(duration, timerWidth, timerHeight, timerLeft, timerTop, sliderColor, SlideWidth());
        }

        #region Body
        private void AddTimerBody(float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, int timerBodyColor)
        {
            var timerBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                timerLeft, timerTop, timerWidth, timerHeight);
            timerBody.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerBodyId);
            timerBody.Fill.ForeColor.RGB = timerBodyColor;
            timerBody.Line.ForeColor.RGB = timerBodyColor;
        }
        #endregion

        #region Markers
        private void AddMarkers(int duration, float timerWidth, float timerHeight, float timerLeft, float timerTop)
        {
            if (duration <= TimerLabConstants.SecondsInMinute)
            {
                AddSecondsMarker(duration, TimerLabConstants.DefaultDenomination, timerWidth, timerHeight, 
                    timerLeft, timerTop, TimerLabConstants.DefaultMinutesLineMarkerWidth, 
                    TimerLabConstants.DefaultTimeMarkerWidth, TimerLabConstants.DefaultTimeMarkerHeight);
            }
            else
            {
                AddMinutesMarker(duration, TimerLabConstants.DefaultDenomination, timerWidth, timerHeight,
                    timerLeft, timerTop, TimerLabConstants.DefaultMinutesLineMarkerWidth,
                    TimerLabConstants.DefaultTimeMarkerWidth, TimerLabConstants.DefaultTimeMarkerHeight);
            }
        }

        private void AddSecondsMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();
            int currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration) 
            {
                // Add time marker
                var timeMarker = AddTimeMarker(timerLeft, currentMarker, widthPerSec, timerTop, timerHeight,
                    timeMarkerWidth, timeMarkerHeight);
                timeMarkers.Add(timeMarker);

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    var lineMarker = AddLineMarker(timerLeft, currentMarker, widthPerSec, timerTop, timerHeight,
                        TimerLabConstants.DefaultSecondsLineMarkerWidth);
                    lineMarkers.Add(lineMarker);
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
                GroupShapes(TimerLabConstants.TimerLineMarkerId, TimerLabConstants.TimerLineMarkerGroupId);
            }
            if (timeMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerTimeMarkerId, TimerLabConstants.TimerTimeMarkerGroupId);
            }
        }

        private void AddMinutesMarker(int duration, int denomination, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight)
        {
            float widthPerSec = timerWidth / duration;

            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();
            int currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                // Add time markers for start, end and every minute
                if (currentMarker % TimerLabConstants.SecondsInMinute == 0 || currentMarker == duration)
                {
                    // Add time marker
                    var timeMarker = AddTimeMarker(timerLeft, currentMarker, widthPerSec, timerTop, timerHeight,
                        timeMarkerWidth, timeMarkerHeight);
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
                        var lineMarker = AddLineMarker(timerLeft, currentMarker, widthPerSec, timerTop, timerHeight,
                        TimerLabConstants.DefaultMinutesLineMarkerWidth);
                        lineMarkers.Add(lineMarker);
                    }
                    else
                    {
                        var lineMarker = AddLineMarker(timerLeft, currentMarker, widthPerSec, timerTop, timerHeight,
                        TimerLabConstants.DefaultSecondsLineMarkerWidth);
                        lineMarkers.Add(lineMarker);
                    }
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
                GroupShapes(TimerLabConstants.TimerLineMarkerId, TimerLabConstants.TimerLineMarkerGroupId);
            }
            if (timeMarkers.Count > 1)
            {
                GroupShapes(TimerLabConstants.TimerTimeMarkerId, TimerLabConstants.TimerTimeMarkerGroupId);
            }
        }

        private Shape AddTimeMarker(float timerLeft, int currentMarker, float widthPerSec, float timerTop, 
            float timerHeight, float timeMarkerWidth, float timeMarkerHeight)
        {
            var timeMarker = this.GetCurrentSlide().Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                (timerLeft + (currentMarker * widthPerSec)) - (timeMarkerWidth / 2), timerTop + timerHeight,
                timeMarkerWidth, timeMarkerHeight);
            timeMarker.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerTimeMarkerId);
            timeMarker.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
            timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
            timeMarker.TextFrame.TextRange.Font.Color.RGB = TimeMarkerColor();
            timeMarker.TextFrame.TextRange.Text = currentMarker.ToString();
            return timeMarker;
        }
        
        private Shape AddLineMarker(float timerLeft, int currentMarker, float widthPerSec, float timerTop, 
            float timerHeight, float lineWeight)
        {
            float beginX = timerLeft + (currentMarker * widthPerSec);
            float beginY = timerTop;
            float endX = timerLeft + (currentMarker * widthPerSec);
            float endY = timerTop + timerHeight;
            var lineMarker = this.GetCurrentSlide().Shapes.AddLine(beginX, beginY, endX, endY);
            lineMarker.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerLineMarkerId);
            lineMarker.Line.Weight = lineWeight;
            return lineMarker;
        }
        #endregion

        #region Slider
        private void AddSlider(int duration, float timerWidth, float timerHeight, float timerLeft, 
            float timerTop, int sliderColor, float slideWidth)
        {
            var sliderHead = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                timerLeft - (TimerLabConstants.DefaultSliderHeadSize / 2), timerTop - (TimerLabConstants.DefaultSliderHeadSize / 2),
                TimerLabConstants.DefaultSliderHeadSize, TimerLabConstants.DefaultSliderHeadSize);
            sliderHead.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerSliderHeadId);
            sliderHead.Rotation = TimerLabConstants.Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = TimerLabConstants.TransparencyTranparent;

            var sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                timerLeft - (TimerLabConstants.DefaultSliderBodyWidth / 2), timerTop, TimerLabConstants.DefaultSliderBodyWidth, timerHeight);
            sliderBody.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerSliderBodyId);
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
        #endregion
        #endregion

        #region Controls
        #region NumericUpDown Control
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

            if (HasTimer())
            {
                ReformMissingComponents();
                UpdateMarkers();
                UpdateSliderDuration();
                AdjustZOrder();
            }
        }
        #endregion

        #region Width Control
        private void WidthSlider_Loaded(object sender, RoutedEventArgs e)
        {
            WidthSlider.Minimum = TimerLabConstants.MinTimerWidth;
            WidthSlider.Maximum = SlideWidth();
            WidthSlider.Value = TimerLabConstants.DefaultTimerWidth;   
        }

        private void WidthSlider_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (HasTimer())
            {
                ReformMissingComponents();
            }
        }

        private void WidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // update text box value
            float value = (float)WidthSlider.Value;
            WidthTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            if (HasTimer())
            {
                var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
                float increment = value - timerBody.Width;

                if (increment == 0)
                {
                    UpdateMarkers();
                }
                else
                {
                    var lineMarkerGroup = GetLineMarkerGroup();
                    var timeMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
                    timerBody.Select();
                    if (lineMarkerGroup != null)
                    {
                        lineMarkerGroup.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
                    }
                    timeMarkerGroup.Select(Microsoft.Office.Core.MsoTriState.msoFalse);

                    var tempGroup = this.GetCurrentSelection().ShapeRange.Group();
                    tempGroup.Left = NewPosition(tempGroup.Left, increment);
                    tempGroup.Width = tempGroup.Width + increment;

                    var ungroupedShapes = tempGroup.Ungroup();
                    timerBody = GetShapeByName(TimerLabConstants.TimerBodyId, ungroupedShapes);
                }

                var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
                var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);
                sliderBody.Left = NewPosition(timerBody.Left, sliderBody.Width);
                sliderHead.Left = NewPosition(timerBody.Left, sliderHead.Width);

                // update animation
                foreach (PowerPoint.Effect effect in this.GetCurrentSlide().TimeLine.MainSequence)
                {
                    if (effect.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectPathRight)
                    {
                        if (effect.Shape.Tags[TimerLabConstants.ShapeId].Equals(TimerLabConstants.TimerSliderBodyId) ||
                            effect.Shape.Tags[TimerLabConstants.ShapeId].Equals(TimerLabConstants.TimerSliderHeadId))
                        {
                            float end = timerBody.Width / SlideWidth();
                            effect.Behaviors[1].MotionEffect.Path = "M 0 0 L " + end + " 0 E";
                        }
                    }
                }
                AdjustZOrder();
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
            else if (value > SlideWidth())
            {
                value = (int)SlideWidth();
            }
            WidthTextBox.Text = value.ToString();
            WidthSlider.Value = value;

            // update timer dimensions
            if (HasTimer())
            {
                ReformMissingComponents();
            }
        }
        #endregion

        #region Height Control
        private void HeightSlider_Loaded(object sender, RoutedEventArgs e)
        {
            HeightSlider.Minimum = TimerLabConstants.MinTimerHeight;
            HeightSlider.Maximum = SlideHeight();
            HeightSlider.Value = TimerLabConstants.DefaultTimerHeight;
        }

        private void HeightSlider_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // update timer dimensions
            if (HasTimer())
            {
                ReformMissingComponents();
            }
        }

        private void HeightSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // update text box value
            float value = (float)HeightSlider.Value;
            HeightTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            if (HasTimer())
            {
                var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);

                float increment = value - timerBody.Height;
                timerBody.Top = NewPosition(timerBody.Top, increment);
                timerBody.Height = timerBody.Height + increment;

                var lineMarkerGroup = GetLineMarkerGroup();
                if (lineMarkerGroup != null)
                {
                    lineMarkerGroup.Top = NewPosition(lineMarkerGroup.Top, increment);
                    lineMarkerGroup.Height = lineMarkerGroup.Height + increment;
                }

                var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);
                sliderBody.Top = NewPosition(sliderBody.Top, increment);
                sliderBody.Height = sliderBody.Height + increment;

                var timeMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
                timeMarkerGroup.Top = NewPosition(timeMarkerGroup.Top, -increment);

                var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
                sliderHead.Top = NewPosition(sliderHead.Top, increment);

                AdjustZOrder();
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
            else if (value > SlideHeight())
            {
                value = (int)SlideHeight();
            }
            HeightTextBox.Text = value.ToString();
            HeightSlider.Value = value;

            // update timer dimensions
            if (HasTimer())
            {
                ReformMissingComponents();
            }
        }
        #endregion
        #endregion

        #region Timer Helper
        private void UpdateMarkers()
        {
            // remove current marker
            Shape lineMarkerGroup = GetLineMarkerGroup();
            if (lineMarkerGroup != null)
            {
                lineMarkerGroup.Delete();
            }
            Shape timerMarkeGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
            timerMarkeGroup.Delete();

            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);

            // add new markers
            AddMarkers(Duration(), timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top);
        }

        private void UpdateSlider()
        {
            // remove current Slider
            Shape sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
            sliderHead.Delete();
            Shape sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);
            sliderBody.Delete();

            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);

            // add new slider
            AddSlider(Duration(), timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top, 
                SliderColor(), SlideWidth());
        }

        private void UpdateSliderDuration()
        {
            foreach (PowerPoint.Effect effect in this.GetCurrentSlide().TimeLine.MainSequence)
            {
                if (effect.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectPathRight)
                {
                    if (effect.Shape.Tags[TimerLabConstants.ShapeId].Equals(TimerLabConstants.TimerSliderBodyId) ||
                        effect.Shape.Tags[TimerLabConstants.ShapeId].Equals(TimerLabConstants.TimerSliderHeadId))
                    {
                        effect.Timing.Duration = Duration();
                    }
                }
            }
        }

        private void AdjustZOrder()
        {
            //Adjust z-order
            var lineMarkerGroup = GetLineMarkerGroup();
            if (lineMarkerGroup != null)
            {
                lineMarkerGroup.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            }
            var timerMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
            timerMarkerGroup.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
            sliderHead.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);
            sliderBody.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
        }

        private void ReformMissingComponents()
        {
            ReformTimerBodyIfMissing();
            ReformMarkersIfMissing();
            ReformSliderIfMissing();
            UpdateMarkers();
            UpdateSlider();
            AdjustZOrder();
        }

        private void ReformTimerBodyIfMissing()
        {
            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
            if (timerBody == null)
            {
                AddTimerBody(TimerWidth(), TimerHeight(), DefaultTimerLeft(SlideWidth(), TimerWidth()),
                    DefaultTimerTop(SlideHeight(), TimerHeight()), TimerBodyColor());
            }
        }

        private void ReformMarkersIfMissing()
        {
            var lineMarkerGroup = GetLineMarkerGroup();
            var timerMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
            if (lineMarkerGroup == null && timerMarkerGroup == null)
            {
                AddMarkers(Duration(), timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top);
            }
            else if (lineMarkerGroup != null && timerMarkerGroup == null)
            {
                lineMarkerGroup.Delete();
                AddMarkers(Duration(), timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top);
            }
            else if (lineMarkerGroup == null && timerMarkerGroup != null)
            {
                timerMarkerGroup.Delete();
                AddMarkers(Duration(), timerBody.Width, timerBody.Height, timerBody.Left, timerBody.Top);
            }
            else
            {
                // Do nothing
            }
        }

        private void ReformSliderIfMissing()
        {
            var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
            var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);
            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
            if (sliderHead == null && sliderBody == null)
            {
                AddSlider(Duration(), timerBody.Width, timerBody.Height, timerBody.Left,
                    timerBody.Top, SliderColor(), SlideWidth());
            }
            else if (sliderHead != null && sliderBody == null)
            {
                sliderHead.Delete();
                AddSlider(Duration(), timerBody.Width, timerBody.Height, timerBody.Left,
                    timerBody.Top, SliderColor(), SlideWidth());
            }
            else if (sliderHead == null && sliderBody != null)
            {
                sliderBody.Delete();
                AddSlider(Duration(), timerBody.Width, timerBody.Height, timerBody.Left,
                    timerBody.Top, SliderColor(), SlideWidth());
            }
            else
            {
                // Do nothing
            }
        }

        private float NewPosition(float originalPosition, float objectSize)
        {
            return originalPosition - objectSize / 2;
        }

        private Shape GetLineMarkerGroup()
        {
            var lineMarkerGroup = GetShapeByName(TimerLabConstants.TimerLineMarkerGroupId);
            if (lineMarkerGroup == null)
            {
                lineMarkerGroup = GetShapeByName(TimerLabConstants.TimerLineMarkerId);
            }
            return lineMarkerGroup;
        }

        private Shape GetShapeByName(string name)
        {
            return GetShapeByName(name, this.GetCurrentSlide().Shapes);
        }

        private Shape GetShapeByName(string name, dynamic shapes)
        {
            foreach (Shape shape in shapes)
            {
                if (shape.Tags[TimerLabConstants.ShapeId].Equals(name))
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
                if (shape.Tags[TimerLabConstants.ShapeId].Equals(shapeName))
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
            group.Tags.Add(TimerLabConstants.ShapeId, groupName);
        }
        #endregion

        #region Validation Helper
        private bool HasTimer()
        {
            var timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
            var lineMarkerGroup = GetLineMarkerGroup();
            var timerMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
            var sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
            var sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);

            if ((timerBody == null) && (lineMarkerGroup == null) &&
                (timerMarkerGroup == null) && (sliderHead == null) && (sliderBody == null))
            {
                return false;
            }
            return true;
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
