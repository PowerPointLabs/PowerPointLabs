using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ColorsLab
{
    /// <summary>
    /// Interaction logic for TimerLabPaneWPF.xaml
    /// </summary>
    public partial class ColorsLabPaneWPF : UserControl
    {
        Shape timerBody = null;
        Shape lineMarkerGroup = null;
        Shape timeMarkerGroup = null;
        Shape sliderHead = null;
        Shape sliderBody = null;

        public ColorsLabPaneWPF()
        {
            InitializeComponent();
        }
        
        #region Timer Properties
        private int Duration()
        {
            int duration = ColorsLabConstants.SecondsInMinute;
            if (DurationTextBox.Value != null)
            {
                double value = Math.Round(DurationTextBox.Value.Value, 2);
                int minutes = (int)value;
                int seconds = (int)(Math.Round(value - minutes, 2) * 100);
                duration = (minutes * ColorsLabConstants.SecondsInMinute) + seconds;
            }
            return duration;
        }

        private bool Countdown()
        {
            bool isCountdown = ColorsLabConstants.DefaultCountdownSetting;
            if (CountdownCheckBox.IsChecked.HasValue)
            {
                isCountdown = CountdownCheckBox.IsChecked.Value;
            }
            return isCountdown;
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
            return GraphicsUtil.PackRgbInt(68, 84, 106);
        }

        private int SliderColor()
        {
            return GraphicsUtil.PackRgbInt(247, 150, 70);
        }

        private int TimeMarkerColor()
        {
            return GraphicsUtil.PackRgbInt(90, 90, 90);
        }

        private int LineMarkerColor()
        {
            return GraphicsUtil.PackRgbInt(68, 114, 196);
        }
        #endregion

        #region Timer Creation
        private void CreateBlocksTimer(int duration, float timerWidth, float timerHeight, float timerLeft, float timerTop, bool isCountdown)
        {
            AddTimerBody(timerWidth, timerHeight, timerLeft, timerTop, TimerBodyColor());
            AddMarkers(duration, timerWidth, timerHeight, TimeMarkerColor(), LineMarkerColor(), isCountdown);
            AddSlider(duration, timerWidth, timerHeight, SliderColor(), SlideWidth());
        }

        #region Body
        private void AddTimerBody(float timerWidth, float timerHeight, float timerLeft, float timerTop, int timerBodyColor)
        {
            timerBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                timerLeft, timerTop, timerWidth, timerHeight);
            timerBody.Name = ColorsLabConstants.TimerBodyId;
            timerBody.Tags.Add(ColorsLabConstants.ShapeId, ColorsLabConstants.TimerBodyId);
            timerBody.Fill.ForeColor.RGB = timerBodyColor;
            timerBody.Line.ForeColor.RGB = timerBodyColor;
        }
        #endregion

        #region Markers
        private void AddMarkers(int duration, float timerWidth, float timerHeight, int timeMarkerColor, int lineMarkerColor, bool isCountdown)
        {
            if (duration <= ColorsLabConstants.SecondsInMinute)
            {
                AddSecondsMarker(duration, ColorsLabConstants.DefaultDenomination, timerWidth, timerHeight, 
                                ColorsLabConstants.DefaultMinutesLineMarkerWidth, 
                                ColorsLabConstants.DefaultTimeMarkerWidth, ColorsLabConstants.DefaultTimeMarkerHeight,
                                timeMarkerColor, lineMarkerColor, isCountdown);
            }
            else
            {
                AddMinutesMarker(duration, ColorsLabConstants.DefaultDenomination, timerWidth, timerHeight,
                                ColorsLabConstants.DefaultMinutesLineMarkerWidth,
                                ColorsLabConstants.DefaultTimeMarkerWidth, ColorsLabConstants.DefaultTimeMarkerHeight,
                                timeMarkerColor, lineMarkerColor, isCountdown);
            }
            UpdateMarkerPosition();
        }

        private void AddSecondsMarker(int duration, int denomination, float timerWidth, float timerHeight, 
                                    float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight,
                                    int timeMarkerColor, int lineMarkerColor, bool isCountdown)
        {
            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();

            float widthPerSec = timerWidth / duration;
            int currentMarker = ColorsLabConstants.StartTime;
            while (currentMarker <= duration) 
            {
                // Get the marker text to be printed
                String markerText = isCountdown ? (duration - currentMarker).ToString() : currentMarker.ToString();

                // Add time marker
                Shape timeMarker = AddTimeMarker(currentMarker, widthPerSec, timerHeight, timeMarkerWidth, timeMarkerHeight, timeMarkerColor, markerText);
                timeMarkers.Add(timeMarker);

                // Add line marker if it is not the start or end
                if (currentMarker != ColorsLabConstants.StartTime && currentMarker != duration)
                {
                    Shape lineMarker = AddLineMarker(currentMarker, widthPerSec, timerHeight, 
                                                    ColorsLabConstants.DefaultSecondsLineMarkerWidth, lineMarkerColor);
                    lineMarkers.Add(lineMarker);
                }

                if (currentMarker >= duration)
                {
                    break;
                }

                currentMarker = GetNextMarkerPosition(currentMarker, duration, denomination, isCountdown);
            }

            lineMarkerGroup = null;
            if (lineMarkers.Count == 1)
            {
                lineMarkerGroup = lineMarkers[0];
            }
            else if (lineMarkers.Count > 1)
            {
                lineMarkerGroup = GroupShapes(ColorsLabConstants.TimerLineMarkerId, ColorsLabConstants.TimerLineMarkerGroupId);
            }
            timeMarkerGroup = GroupShapes(ColorsLabConstants.TimerTimeMarkerId, ColorsLabConstants.TimerTimeMarkerGroupId);
        }

        private void AddMinutesMarker(int duration, int denomination, float timerWidth, float timerHeight, 
                                    float lineMarkerWidth, float timeMarkerWidth, float timeMarkerHeight, 
                                    int timeMarkerColor, int lineMarkerColor, bool isCountdown)
        {
            List<Shape> lineMarkers = new List<Shape>();
            List<Shape> timeMarkers = new List<Shape>();

            float widthPerSec = timerWidth / duration;
            int currentMarker = ColorsLabConstants.StartTime;
            while (currentMarker <= duration)
            {
                bool isStart = currentMarker == 0;
                bool isMinuteMark = isCountdown ? (duration - currentMarker) % ColorsLabConstants.SecondsInMinute == 0 : currentMarker % ColorsLabConstants.SecondsInMinute == 0;
                bool isEnd = currentMarker == duration;

                // Add time markers for start, every minute, and end
                if (isStart || isMinuteMark || isEnd)
                {
                    // Add time marker
                    Shape timeMarker = AddMinuteTimeMarker(duration, currentMarker, widthPerSec, timerHeight, timeMarkerWidth, timeMarkerHeight, timeMarkerColor, isCountdown);
                    timeMarkers.Add(timeMarker);
                }

                // Add line marker if it is not the start or end
                if (currentMarker != ColorsLabConstants.StartTime && currentMarker != duration)
                {
                    // Thicken the line if it is a minute marker
                    Shape lineMarker = AddMinuteLineMarker(duration, currentMarker, widthPerSec, timerHeight, lineMarkerColor, isCountdown);
                    lineMarkers.Add(lineMarker);
                }

                if (currentMarker >= duration)
                {
                    break;
                }

                currentMarker = GetNextMarkerPosition(currentMarker, duration, denomination, isCountdown);
            }

            lineMarkerGroup = GroupShapes(ColorsLabConstants.TimerLineMarkerId, ColorsLabConstants.TimerLineMarkerGroupId);
            timeMarkerGroup = GroupShapes(ColorsLabConstants.TimerTimeMarkerId, ColorsLabConstants.TimerTimeMarkerGroupId);
        }


        private int GetNextMarkerPosition(int currentMarker, int duration, int denomination, bool isCountdown)
        {
            // If it's Countdown Timer and we are at the start, take into account specified durations that are not multiple of denomination
            if (isCountdown && currentMarker == 0 && duration % denomination != 0)
            {
                return duration % denomination;
            }
            else
            {
                return Math.Min(currentMarker + denomination, duration);
            }
        }

        private Shape AddMinuteTimeMarker(int duration, int currentMarker, float widthPerSec, float timerHeight,
                                    float timeMarkerWidth, float timeMarkerHeight, int timeMarkerColor, bool isCountdown)
        {
            // Get the marker text to be printed
            int remainingDuration = duration - currentMarker;
            String markerText = isCountdown ? remainingDuration.ToString() : currentMarker.ToString();

            Shape timeMarker = AddTimeMarker(currentMarker, widthPerSec, timerHeight, timeMarkerWidth, timeMarkerHeight, timeMarkerColor, markerText);

            if (!isCountdown)
            {
                int remainingSeconds = currentMarker % ColorsLabConstants.SecondsInMinute;
                if (currentMarker == duration && remainingSeconds != 0)
                {
                    timeMarker.TextFrame.TextRange.Text = (currentMarker / ColorsLabConstants.SecondsInMinute).ToString() +
                                                            "." + remainingSeconds.ToString("D2");
                }
                else
                {
                    timeMarker.TextFrame.TextRange.Text = (currentMarker / ColorsLabConstants.SecondsInMinute).ToString();
                }
            }
            else
            {
                int leftoverSeconds = remainingDuration % ColorsLabConstants.SecondsInMinute;
                if (currentMarker == 0 && leftoverSeconds != 0)
                {
                    timeMarker.TextFrame.TextRange.Text = (remainingDuration / ColorsLabConstants.SecondsInMinute).ToString() +
                                                            "." + leftoverSeconds.ToString("D2");
                }
                else
                {
                    timeMarker.TextFrame.TextRange.Text = (remainingDuration / ColorsLabConstants.SecondsInMinute).ToString();
                }
            }

            return timeMarker;
        }

        private Shape AddMinuteLineMarker(int duration, int currentMarker, float widthPerSec, float timerHeight, 
                                          int lineMarkerColor, bool isCountdown)
        {
            bool isMinuteMarker = isCountdown ? ((duration - currentMarker) % ColorsLabConstants.SecondsInMinute == 0) :
                                                         (currentMarker % ColorsLabConstants.SecondsInMinute == 0);
            float markerLineWeight = isMinuteMarker ? ColorsLabConstants.DefaultMinutesLineMarkerWidth :
                                                        ColorsLabConstants.DefaultSecondsLineMarkerWidth;
            Shape lineMarker = AddLineMarker(currentMarker, widthPerSec, timerHeight, markerLineWeight, lineMarkerColor);
            return lineMarker;
        }

        private Shape AddTimeMarker(int currentMarker, float widthPerSec, float timerHeight, 
                                    float timeMarkerWidth, float timeMarkerHeight, int timeMarkerColor, string markerText)
        {
            Shape timeMarker = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                    currentMarker * widthPerSec, 0, 
                                                                    timeMarkerWidth, timeMarkerHeight);
            timeMarker.Name = ColorsLabConstants.TimerTimeMarkerId + markerText;
            timeMarker.Tags.Add(ColorsLabConstants.ShapeId, ColorsLabConstants.TimerTimeMarkerId);
            timeMarker.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            timeMarker.Fill.Transparency = ColorsLabConstants.TransparencyTranparent;
            timeMarker.Line.Transparency = ColorsLabConstants.TransparencyTranparent;
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
            lineMarker.Name = ColorsLabConstants.TimerLineMarkerId + markerText;
            lineMarker.Tags.Add(ColorsLabConstants.ShapeId, ColorsLabConstants.TimerLineMarkerId);
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
                                                                ColorsLabConstants.DefaultSliderHeadSize, 
                                                                ColorsLabConstants.DefaultSliderHeadSize);
            sliderHead.Name = ColorsLabConstants.TimerSliderHeadId;
            sliderHead.Tags.Add(ColorsLabConstants.ShapeId, ColorsLabConstants.TimerSliderHeadId);
            sliderHead.Rotation = ColorsLabConstants.Rotate180Degrees;
            sliderHead.Fill.ForeColor.RGB = sliderColor;
            sliderHead.Line.Transparency = ColorsLabConstants.TransparencyTranparent;

            sliderBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                                                                0, 0, 
                                                                ColorsLabConstants.DefaultSliderBodyWidth, 
                                                                timerHeight);
            sliderBody.Name = ColorsLabConstants.TimerSliderBodyId;
            sliderBody.Tags.Add(ColorsLabConstants.ShapeId, ColorsLabConstants.TimerSliderBodyId);
            sliderBody.Fill.ForeColor.RGB = sliderColor;
            sliderBody.Line.Transparency = ColorsLabConstants.TransparencyTranparent;

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
            sliderEndEffect.Timing.Duration = ColorsLabConstants.ColorChangeDuration;
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

                WidthTextBox.Text = Math.Round(timerBody.Width).ToString();
                HeightTextBox.Text = Math.Round(timerBody.Height).ToString();

                ShowErrorMessageBox(ColorsLabConstants.ErrorMessageOneTimerOnly);
            }
            else
            {
                // Properties
                int duration = Duration();
                bool isCountdown = Countdown();
                float timerWidth = TimerWidth();
                float timerHeight = TimerHeight();

                // Position
                float timerLeft = DefaultTimerLeft(SlideWidth(), timerWidth);
                float timerTop = DefaultTimerTop(SlideHeight(), timerHeight);

                CreateBlocksTimer(duration, timerWidth, timerHeight, timerLeft, timerTop, isCountdown);
            }
        }
        #endregion

        #region Duration Control
        private void DurationTextBox_ValueDecremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                DurationTextBox.Value = ColorsLabConstants.DefaultDisplayDuration;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == ColorsLabConstants.FractionalDecrementLowerBound)
            {
                DurationTextBox.Value = (integerPart - 1) + ColorsLabConstants.FractionalDecrementOffset;
            }
        }

        private void DurationTextBox_ValueIncremented(object sender, 
            MahApps.Metro.Controls.NumericUpDownChangedRoutedEventArgs args)
        {
            if (DurationTextBox.Value == null)
            {
                DurationTextBox.Value = ColorsLabConstants.DefaultDisplayDuration;
            }

            double value = Math.Round(DurationTextBox.Value.Value, 2);
            int integerPart = (int)value;
            double fractionalPart = value - integerPart;

            if (Math.Round(fractionalPart, 2) == ColorsLabConstants.FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + ColorsLabConstants.FractionalIncrementOffset;
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

            if (Math.Round(fractionalPart, 2) > ColorsLabConstants.FractionalIncrementUpperBound)
            {
                DurationTextBox.Value = integerPart + 1;
            }

            if (FindTimer())
            {
                ReformMissingComponents();
                RecreateMarkers();
                AdjustZOrder();
                UpdateSliderPosition();
                UpdateSliderAnimationDuration();
            }
        }
        #endregion

        #region Countdown Control

        private void CountdownCheckBox_StateChanged(object sender, RoutedEventArgs e)
        {
            // CountdownCheckBox.isChecked can return null if checkbox is in indeterminate state in a 3-state checkbox (checked, unchecked, indeterminate)
            // In this application, the checkbox is only 2-state, but we guard against this because IsChecked returns a nullable boolean (bool?)
            if (CountdownCheckBox.IsChecked == null)
            {
                return;
            }

            if (FindTimer())
            {
                ReformMissingComponents();
                RecreateMarkers();
                AdjustZOrder();
                UpdateSliderPosition();
                UpdateSliderAnimationDuration();
            }
        }

        #endregion

        #region Width Control
        private void WidthSlider_Loaded(object sender, RoutedEventArgs e)
        {
            WidthSlider.Minimum = ColorsLabConstants.MinTimerWidth;
            WidthSlider.Maximum = SlideWidth();
            WidthSlider.Value = ColorsLabConstants.DefaultTimerWidth;   
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
            WidthTextBox.Text = ColorsLabConstants.DefaultTimerWidth.ToString();
        }

        private void WidthTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumbersOnly(e.Text);
        }

        private void WidthTextBox_TextBoxPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsNumbersOnly(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        private void WidthTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(WidthTextBox.Text))
            {
                WidthTextBox.Text = ((int)WidthSlider.Value).ToString();
                return;
            }

            int value = Convert.ToInt32(WidthTextBox.Text);
            value = Math.Max(value, (int)ColorsLabConstants.MinTimerWidth);
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
            HeightSlider.Minimum = ColorsLabConstants.MinTimerHeight;
            HeightSlider.Maximum = SlideHeight();
            HeightSlider.Value = ColorsLabConstants.DefaultTimerHeight;
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
            HeightTextBox.Text = ColorsLabConstants.DefaultTimerHeight.ToString();
        }

        private void HeightTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumbersOnly(e.Text);
        }

        private void HeightTextBox_TextBoxPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsNumbersOnly(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        private void HeightTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(HeightTextBox.Text))
            {
                HeightTextBox.Text = ((int)HeightSlider.Value).ToString();
                return;
            }

            int value = Convert.ToInt32(HeightTextBox.Text);
            value = Math.Max(value, (int)ColorsLabConstants.MinTimerHeight);
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
            AddMarkers(Duration(), timerBody.Width, timerBody.Height, timeMarkerColor, lineMarkerColor, Countdown());
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
                float lineSpacing = ColorsLabConstants.DefaultDenomination * widthPerSec;
                int numOfLineMarkers = (int)(Math.Ceiling((double)Duration() / ColorsLabConstants.DefaultDenomination)) - 2;
                lineMarkerGroup.Left = timerBody.Left + lineSpacing;
                lineMarkerGroup.Width = numOfLineMarkers * lineSpacing;

                // Countdown timers have inconsistent starting points, espeically when duration of the timer is not a multiple of the denomination (10 sec)
                // So we need to take this into account by calculating the required space and resetting the lineMarkerGroup
                // This is unlike the default timer where the starting offset is always the same (1 lineSpacing from left)
                if (Countdown())
                {
                    float requiredSpaceFromLeft = timerBody.Width - lineSpacing - lineMarkerGroup.Width;
                    lineMarkerGroup.Left = timerBody.Left + requiredSpaceFromLeft;
                }
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
            sliderHead.Left = timerBody.Left - (ColorsLabConstants.DefaultSliderHeadSize / 2);
            sliderBody.Left = timerBody.Left - (ColorsLabConstants.DefaultSliderBodyWidth / 2);
            UpdateSliderAnimationPath();
        }

        private void UpdateSliderPositionY()
        {
            sliderHead.Top = timerBody.Top - (ColorsLabConstants.DefaultSliderHeadSize / 2);
            sliderBody.Top = timerBody.Top;
            sliderBody.Height = timerBody.Height;
        }

        private void UpdateSliderAnimationDuration()
        {
            foreach (PowerPoint.Effect effect in this.GetCurrentSlide().TimeLine.MainSequence)
            {
                if (effect.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectPathRight)
                {
                    if (effect.Shape.Tags[ColorsLabConstants.ShapeId].Equals(ColorsLabConstants.TimerSliderBodyId) ||
                        effect.Shape.Tags[ColorsLabConstants.ShapeId].Equals(ColorsLabConstants.TimerSliderHeadId))
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
                    if (effect.Shape.Tags[ColorsLabConstants.ShapeId].Equals(ColorsLabConstants.TimerSliderBodyId) ||
                        effect.Shape.Tags[ColorsLabConstants.ShapeId].Equals(ColorsLabConstants.TimerSliderHeadId))
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
            Shape result = GetShapeByName(ColorsLabConstants.TimerLineMarkerGroupId);
            if (result == null)
            {
                result = GetShapeByName(ColorsLabConstants.TimerLineMarkerId);
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
                if (shape.Tags[ColorsLabConstants.ShapeId].Equals(name))
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
                if (shape.Tags[ColorsLabConstants.ShapeId].Equals(shapeName))
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
            group.Tags.Add(ColorsLabConstants.ShapeId, groupName);
            return group;
        }
        #endregion

        #region Validation Helper
        private bool FindTimer()
        {
            timerBody = GetShapeByName(ColorsLabConstants.TimerBodyId);
            lineMarkerGroup = GetLineMarkerGroup();
            timeMarkerGroup = GetShapeByName(ColorsLabConstants.TimerTimeMarkerGroupId);
            sliderHead = GetShapeByName(ColorsLabConstants.TimerSliderHeadId);
            sliderBody = GetShapeByName(ColorsLabConstants.TimerSliderBodyId);

            if ((timerBody == null) && (lineMarkerGroup == null) &&
                (timeMarkerGroup == null) && (sliderHead == null) && (sliderBody == null))
            {
                return false;
            }
            return true;
        }

        private bool IsNumbersOnly(string text)
        {
            Regex regex = new Regex("[0-9]+");
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
            MessageBox.Show(content, TextCollection.CommonText.ErrorTitle);
        }

        #endregion
    }
}
