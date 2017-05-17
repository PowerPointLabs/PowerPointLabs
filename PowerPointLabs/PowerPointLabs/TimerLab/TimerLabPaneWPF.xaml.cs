using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.TimerLab
{
    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class TimerLabPaneWPF : UserControl
    {
        Shape timerBody = null;
        Shape lineMarkerGroup = null;
        Shape timeMarkerGroup = null;
        Shape sliderHead = null;
        Shape sliderBody = null;

        public TimerLabPaneWPF()
        {
            InitializeComponent();
        }
        
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
            return Graphics.PackRgbInt(68, 84, 106);
        }

        private int SliderColor()
        {
            return Graphics.PackRgbInt(247, 150, 70);
        }

        private int TimeMarkerColor()
        {
            return Graphics.PackRgbInt(90, 90, 90);
        }

        private int LineMarkerColor()
        {
            return Graphics.PackRgbInt(68, 114, 196);
        }
        #endregion

        #region Timer Creation
        private void CreateBlocksTimer(int duration, float timerWidth, float timerHeight, float timerLeft, float timerTop)
        {
            AddTimerBody(timerWidth, timerHeight, timerLeft, timerTop, TimerBodyColor());
            AddMarkers(duration, timerWidth, timerHeight, TimeMarkerColor(), LineMarkerColor());
            AddSlider(duration, timerWidth, timerHeight, SliderColor(), SlideWidth());
        }

        #region Body
        private void AddTimerBody(float timerWidth, float timerHeight, float timerLeft, float timerTop, int timerBodyColor)
        {
            timerBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                timerLeft, timerTop, timerWidth, timerHeight);
            timerBody.Name = TimerLabConstants.TimerBodyId;
            timerBody.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerBodyId);
            timerBody.Fill.ForeColor.RGB = timerBodyColor;
            timerBody.Line.ForeColor.RGB = timerBodyColor;
        }
        #endregion

        #region Markers
        private void AddMarkers(int duration, float timerWidth, float timerHeight, int timeMarkerColor, int lineMarkerColor)
        {
            if (duration <= TimerLabConstants.SecondsInMinute)
            {
                AddSecondsMarker(duration, TimerLabConstants.DefaultDenomination, timerWidth, timerHeight, 
                                TimerLabConstants.DefaultMinutesLineMarkerWidth, 
                                TimerLabConstants.DefaultTimeMarkerWidth, TimerLabConstants.DefaultTimeMarkerHeight,
                                timeMarkerColor, lineMarkerColor);
            }
            else
            {
                AddMinutesMarker(duration, TimerLabConstants.DefaultDenomination, timerWidth, timerHeight,
                                TimerLabConstants.DefaultMinutesLineMarkerWidth,
                                TimerLabConstants.DefaultTimeMarkerWidth, TimerLabConstants.DefaultTimeMarkerHeight,
                                timeMarkerColor, lineMarkerColor);
            }
            UpdateMarkerPosition();
        }

        private void AddSecondsMarker(int duration, int denomination, float timerWidth, float timerHeight, 
                                    float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight,
                                    int timeMarkerColor, int lineMarkerColor)
        {
            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();

            float widthPerSec = timerWidth / duration;
            int currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration) 
            {
                // Add time marker
                Shape timeMarker = AddTimeMarker(currentMarker, widthPerSec, timerHeight, timeMarkerWidth, timeMarkerHeight, timeMarkerColor);
                timeMarkers.Add(timeMarker);

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    Shape lineMarker = AddLineMarker(currentMarker, widthPerSec, timerHeight, 
                                                    TimerLabConstants.DefaultSecondsLineMarkerWidth, lineMarkerColor);
                    lineMarkers.Add(lineMarker);
                }

                if (currentMarker >= duration)
                {
                    break;
                }
                currentMarker = Math.Min(currentMarker + denomination, duration);
            }

            lineMarkerGroup = null;
            if (lineMarkers.Count == 1)
            {
                lineMarkerGroup = lineMarkers[0];
            }
            else if (lineMarkers.Count > 1)
            {
                lineMarkerGroup = GroupShapes(TimerLabConstants.TimerLineMarkerId, TimerLabConstants.TimerLineMarkerGroupId);
            }
            timeMarkerGroup = GroupShapes(TimerLabConstants.TimerTimeMarkerId, TimerLabConstants.TimerTimeMarkerGroupId);
        }

        private void AddMinutesMarker(int duration, int denomination, float timerWidth, float timerHeight, 
                                    float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight, 
                                    int timeMarkerColor, int lineMarkerColor)
        {
            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();

            float widthPerSec = timerWidth / duration;
            int currentMarker = TimerLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                // Add time markers for start, every minute, and end
                if (currentMarker % TimerLabConstants.SecondsInMinute == 0 || currentMarker == duration)
                {
                    // Add time marker
                    Shape timeMarker = AddTimeMarker(currentMarker, widthPerSec, timerHeight, timeMarkerWidth, timeMarkerHeight, timeMarkerColor);

                    int remainingSeconds = currentMarker % TimerLabConstants.SecondsInMinute;
                    if (currentMarker == duration && remainingSeconds != 0)
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / TimerLabConstants.SecondsInMinute).ToString() + 
                                                                "." + remainingSeconds.ToString("D2");
                    }
                    else
                    {
                        timeMarker.TextFrame.TextRange.Text = (currentMarker / TimerLabConstants.SecondsInMinute).ToString();
                    }

                    timeMarkers.Add(timeMarker);
                }

                // Add line marker if it is not the start or end
                if (currentMarker != TimerLabConstants.StartTime && currentMarker != duration)
                {
                    //Thicken the line if it is a minute marker
                    bool isMinuteMarker = (currentMarker % TimerLabConstants.SecondsInMinute == 0);
                    float markerLineWeight = isMinuteMarker ? TimerLabConstants.DefaultMinutesLineMarkerWidth :
                                                                TimerLabConstants.DefaultSecondsLineMarkerWidth;
                    Shape lineMarker = AddLineMarker(currentMarker, widthPerSec, timerHeight, markerLineWeight, lineMarkerColor);
                    lineMarkers.Add(lineMarker);
                }

                if (currentMarker >= duration)
                {
                    break;
                }
                currentMarker = Math.Min(currentMarker + denomination, duration);
            }

            lineMarkerGroup = GroupShapes(TimerLabConstants.TimerLineMarkerId, TimerLabConstants.TimerLineMarkerGroupId);
            timeMarkerGroup = GroupShapes(TimerLabConstants.TimerTimeMarkerId, TimerLabConstants.TimerTimeMarkerGroupId);
        }

        private Shape AddTimeMarker(int currentMarker, float widthPerSec, float timerHeight, 
                                    float timeMarkerWidth, float timeMarkerHeight, int timeMarkerColor)
        {
            string markerText = currentMarker.ToString();
            Shape timeMarker = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                    currentMarker * widthPerSec, 0, 
                                                                    timeMarkerWidth, timeMarkerHeight);
            timeMarker.Name = TimerLabConstants.TimerTimeMarkerId + markerText;
            timeMarker.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerTimeMarkerId);
            timeMarker.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            timeMarker.Fill.Transparency = TimerLabConstants.TransparencyTranparent;
            timeMarker.Line.Transparency = TimerLabConstants.TransparencyTranparent;
            timeMarker.TextFrame.TextRange.Font.Color.RGB = timeMarkerColor;
            timeMarker.TextFrame.TextRange.Text = markerText;
            return timeMarker;
        }
        
        private Shape AddLineMarker(int currentMarker, float widthPerSec, float timerHeight, float lineWeight, int lineMarkerColor)
        {
            string markerText = currentMarker.ToString();
            float beginX = currentMarker * widthPerSec;
            float beginY = 0;
            float endX = currentMarker * widthPerSec;
            float endY = timerHeight;

            Shape lineMarker = this.GetCurrentSlide().Shapes.AddLine(beginX, beginY, endX, endY);
            lineMarker.Name = TimerLabConstants.TimerLineMarkerId + markerText;
            lineMarker.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerLineMarkerId);
            lineMarker.Line.Weight = lineWeight;
            lineMarker.Line.ForeColor.RGB = lineMarkerColor;

            return lineMarker;
        }
        #endregion

        #region Slider
        private void AddSlider(int duration, float timerWidth, float timerHeight, int sliderColor, float slideWidth)
        {
            sliderHead = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                                0, 0, 
                                                                TimerLabConstants.DefaultSliderHeadSize, 
                                                                TimerLabConstants.DefaultSliderHeadSize);
            sliderHead.Name = TimerLabConstants.TimerSliderHeadId;
            sliderHead.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerSliderHeadId);
            sliderHead.Rotation = TimerLabConstants.Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = TimerLabConstants.TransparencyTranparent;

            sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                0, 0, 
                                                                TimerLabConstants.DefaultSliderBodyWidth, 
                                                                timerHeight);
            sliderBody.Name = TimerLabConstants.TimerSliderBodyId;
            sliderBody.Tags.Add(TimerLabConstants.ShapeId, TimerLabConstants.TimerSliderBodyId);
            sliderBody.Fill.ForeColor.RGB = sliderColor;
            sliderBody.Line.Transparency = TimerLabConstants.TransparencyTranparent;

            UpdateSliderPosition();

            // Add slider animations
            AddSliderMotionEffect(sliderHead, duration, timerWidth, slideWidth, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            AddSliderMotionEffect(sliderBody, duration, timerWidth, slideWidth, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            AddSliderEndEffect(sliderHead, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            AddSliderEndEffect(sliderBody, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
        }

        private void AddSliderMotionEffect(Shape shape, int duration, float timerWidth, float slideWidth, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Effect sliderMotionEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(
                    shape,
                    PowerPoint.MsoAnimEffect.msoAnimEffectPathRight, 
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, 
                    trigger);
            PowerPoint.AnimationBehavior motion = sliderMotionEffect.Behaviors[1];
            float end = timerWidth / slideWidth;
            motion.MotionEffect.Path = "M 0 0 L " + end + " 0 E";
            sliderMotionEffect.Timing.Duration = duration;
            sliderMotionEffect.Timing.SmoothStart = Microsoft.Office.Core.MsoTriState.msoFalse;
            sliderMotionEffect.Timing.SmoothEnd = Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        private void AddSliderEndEffect(Shape shape, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Effect sliderEndEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(
                    shape,
                    PowerPoint.MsoAnimEffect.msoAnimEffectDarken, 
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    trigger);
            sliderEndEffect.Timing.Duration = TimerLabConstants.ColorChangeDuration;
        }
        #endregion
        #endregion

        #region Controls
        #region Create Button
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // check if timer is already created
            if (FindTimer())
            {
                ReformMissingComponents();
                UpdateMarkerPosition();
                UpdateSliderPosition();
                ShowErrorMessageBox(TimerLabConstants.ErrorMessageOneTimerOnly);
            }
            else
            {
                // Properties
                int duration = Duration();
                float timerWidth = TimerWidth();
                float timerHeight = TimerHeight();

                // Position
                float timerLeft = DefaultTimerLeft(SlideWidth(), timerWidth);
                float timerTop = DefaultTimerTop(SlideHeight(), timerHeight);

                CreateBlocksTimer(duration, timerWidth, timerHeight, timerLeft, timerTop);
            }
        }
        #endregion

        #region Duration Control
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

            if (FindTimer())
            {
                ReformMissingComponents();
                RecreateMarkers();
                UpdateSliderPosition();
                UpdateSliderAnimationDuration();
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

        private void WidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // update text box value
            float value = (float)WidthSlider.Value;
            WidthTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            if (FindTimer())
            {
                ReformMissingComponents();
                
                float increment = value - timerBody.Width;
                timerBody.Left = NewPosition(timerBody.Left, increment);
                timerBody.Width = timerBody.Width + increment;

                UpdateMarkerPositionX();
                UpdateSliderPositionX();
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
            value = Math.Max(value, (int)TimerLabConstants.MinTimerWidth);
            value = Math.Min(value, (int)SlideWidth());
            WidthTextBox.Text = value.ToString();
            WidthSlider.Value = value;

            // update timer dimensions
            if (FindTimer())
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

        private void HeightSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // update text box value
            float value = (float)HeightSlider.Value;
            HeightTextBox.Text = ((int)value).ToString();

            // update timer dimensions
            if (FindTimer())
            {
                ReformMissingComponents();

                float increment = value - timerBody.Height;
                timerBody.Top = NewPosition(timerBody.Top, increment);
                timerBody.Height = timerBody.Height + increment;

                UpdateMarkerPositionY();
                UpdateSliderPositionY();
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
            value = Math.Max(value, (int)TimerLabConstants.MinTimerHeight);
            value = Math.Min(value, (int)SlideHeight());
            HeightTextBox.Text = value.ToString();
            HeightSlider.Value = value;

            // update timer dimensions
            if (FindTimer())
            {
                ReformMissingComponents();
            }
        }
        #endregion
        #endregion

        #region Timer Helper
        private void ReformMissingComponents()
        {
            bool isTimerBodyRecreated = ReformTimerBodyIfMissing();
            bool isMarkersRecreated = ReformMarkersIfMissing();
            bool isSliderRecreated = ReformSliderIfMissing();

            if (isTimerBodyRecreated || isMarkersRecreated || isSliderRecreated)
            {
                AdjustZOrder();
            }
        }

        private bool ReformTimerBodyIfMissing()
        {
            if (timerBody == null)
            {
                AddTimerBody(TimerWidth(), TimerHeight(), 
                            DefaultTimerLeft(SlideWidth(), TimerWidth()),
                            DefaultTimerTop(SlideHeight(), TimerHeight()), 
                            TimerBodyColor());
                return true;
            }
            timerBody.Rotation = 0;
            return false;
        }

        private bool ReformMarkersIfMissing()
        {
            if (lineMarkerGroup == null || timeMarkerGroup == null)
            {
                RecreateMarkers();
                return true;
            }
            return false;
        }

        private bool ReformSliderIfMissing()
        {
            if (sliderHead == null || sliderBody == null)
            {
                RecreateSlider();
                return true;
            }
            return false;
        }

        private void RecreateMarkers()
        {
            // remove current markers
            int timeMarkerColor = TimeMarkerColor();
            if (timeMarkerGroup != null)
            {
                timeMarkerColor = timeMarkerGroup.TextFrame.TextRange.Font.Color.RGB;
                timeMarkerGroup.Delete();
            }

            int lineMarkerColor = LineMarkerColor();
            if (lineMarkerGroup != null)
            {
                lineMarkerColor = lineMarkerGroup.Line.ForeColor.RGB;
                lineMarkerGroup.Delete();
            }

            // add new markers
            AddMarkers(Duration(), timerBody.Width, timerBody.Height, timeMarkerColor, lineMarkerColor);
            timeMarkerGroup.TextFrame.TextRange.Font.Color.RGB = timeMarkerColor;
        }

        private void RecreateSlider()
        {
            int sliderColor = SliderColor();

            // remove current Slider
            if (sliderHead != null)
            {
                sliderColor = sliderHead.Fill.ForeColor.RGB;
                sliderHead.Delete();
            }
            if (sliderBody != null)
            {
                sliderColor = sliderBody.Fill.ForeColor.RGB;
                sliderBody.Delete();
            }
            
            AddSlider(Duration(), timerBody.Width, timerBody.Height, sliderColor, SlideWidth());
        }

        private void UpdateMarkerPosition()
        {
            UpdateMarkerPositionX();
            UpdateMarkerPositionY();
        }

        private void UpdateMarkerPositionX()
        {
            if (lineMarkerGroup != null)
            {
                float widthPerSec = timerBody.Width / Duration();
                float lineSpacing = TimerLabConstants.DefaultDenomination * widthPerSec;
                int numOfLineMarkers = (int)(Math.Ceiling((double)Duration() / TimerLabConstants.DefaultDenomination)) - 2;
                lineMarkerGroup.Left = timerBody.Left + lineSpacing;
                lineMarkerGroup.Width = numOfLineMarkers * lineSpacing;
            }
            timeMarkerGroup.Left = timerBody.Left;
            timeMarkerGroup.Width = timerBody.Width;
        }

        private void UpdateMarkerPositionY()
        {
            if (lineMarkerGroup != null)
            {
                lineMarkerGroup.Top = timerBody.Top;
                lineMarkerGroup.Height = timerBody.Height;
            }
            timeMarkerGroup.Top = timerBody.Top + timerBody.Height;
        }

        private void UpdateSliderPosition()
        {
            UpdateSliderPositionX();
            UpdateSliderPositionY();
        }

        private void UpdateSliderPositionX()
        {
            sliderHead.Left = timerBody.Left - (TimerLabConstants.DefaultSliderHeadSize / 2);
            sliderBody.Left = timerBody.Left - (TimerLabConstants.DefaultSliderBodyWidth / 2);
            UpdateSliderAnimationPath();
        }

        private void UpdateSliderPositionY()
        {
            sliderHead.Top = timerBody.Top - (TimerLabConstants.DefaultSliderHeadSize / 2);
            sliderBody.Top = timerBody.Top;
            sliderBody.Height = timerBody.Height;
        }

        private void UpdateSliderAnimationDuration()
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

        private void UpdateSliderAnimationPath()
        {
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
        }

        private void AdjustZOrder()
        {
            //Adjust z-order
            if (lineMarkerGroup != null)
            {
                lineMarkerGroup.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            }
            timeMarkerGroup.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            sliderHead.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            sliderBody.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
        }
        #endregion

        #region Shape Helper
        private float NewPosition(float originalPosition, float objectSize)
        {
            return originalPosition - objectSize / 2;
        }

        private Shape GetLineMarkerGroup()
        {
            Shape result = GetShapeByName(TimerLabConstants.TimerLineMarkerGroupId);
            if (result == null)
            {
                result = GetShapeByName(TimerLabConstants.TimerLineMarkerId);
            }
            return result;
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

        private Shape GroupShapes(string shapeName, string groupName)
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
            group.Name = groupName;
            group.Tags.Add(TimerLabConstants.ShapeId, groupName);
            return group;
        }
        #endregion

        #region Validation Helper
        private bool FindTimer()
        {
            timerBody = GetShapeByName(TimerLabConstants.TimerBodyId);
            lineMarkerGroup = GetLineMarkerGroup();
            timeMarkerGroup = GetShapeByName(TimerLabConstants.TimerTimeMarkerGroupId);
            sliderHead = GetShapeByName(TimerLabConstants.TimerSliderHeadId);
            sliderBody = GetShapeByName(TimerLabConstants.TimerSliderBodyId);

            if ((timerBody == null) && (lineMarkerGroup == null) &&
                (timeMarkerGroup == null) && (sliderHead == null) && (sliderBody == null))
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

        #region Error Handling
        private void ShowErrorMessageBox(string content)
        {
            MessageBox.Show(content, "Error");
        }

        #endregion
    }
}
