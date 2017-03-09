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

        private void BlocksTimer_Click(object sender, RoutedEventArgs e)
        {
            // slide dimensions
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            // timer specifications
            var time = 60;
            var timerWidth = 600;
            var timerHeight = 50;
            var lineIndicatorWidth = 1;
            var timeIndicatorWidth = 50;
            var timeIndicatorHeight = 50;
            var denomination = 10;

            var timerStartX = (slideWidth - timerWidth) / 2;
            var timerStartY = (slideHeight - timerHeight) / 2;
            var widthPerSec = timerWidth / time;

            var timerBody = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, timerStartX, timerStartY, timerWidth, timerHeight);
            timerBody.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(106, 84, 68).ToArgb();
            timerBody.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(106, 84, 68).ToArgb();
            timerBody.Name = "timerBody";

            var curIndicator = 0;
            while (curIndicator <= time)
            {
                if (curIndicator != 0 && curIndicator != time)
                {
                    var lineIndicator = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, timerStartX + (curIndicator * widthPerSec), timerStartY, lineIndicatorWidth, timerHeight);
                    lineIndicator.Name = "lineIndicator";
                }
                var timeIndicator = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, (timerStartX + (curIndicator * widthPerSec)) - (timeIndicatorWidth / 2), timerStartY + timerHeight, timeIndicatorWidth, timeIndicatorHeight);
                timeIndicator.Name = "timeIndicator";
                timeIndicator.TextFrame.TextRange.Text = curIndicator.ToString();
                timeIndicator.TextFrame.TextRange.Font.Color.RGB = 0;
                timeIndicator.Fill.Transparency = 1;
                timeIndicator.Line.Transparency = 1;

                curIndicator += denomination;
            }

            var sliderHead = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, timerStartX - (float)7.5, timerStartY - (float)7.5, 20, 20);
            sliderHead.Rotation = 180;
            sliderHead.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(247, 149, 50).ToArgb();
            sliderHead.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(247, 149, 50).ToArgb();
            var slider = this.GetCurrentSlide().Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, timerStartX, timerStartY, 5, timerHeight);
            slider.Name = "slider";
            slider.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(247, 149, 50).ToArgb();
            slider.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(247, 149, 50).ToArgb();
            sliderHead.Select();
            slider.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
            PowerPoint.ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            shapeRange.Group().Select();

            var sliderIndicator = this.GetCurrentSelection().ShapeRange[1];
            
            PowerPoint.Effect sliderMotionEffect = this.GetCurrentSlide().TimeLine.MainSequence.AddEffect(sliderIndicator,
                PowerPoint.MsoAnimEffect.msoAnimEffectPathRight,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            PowerPoint.AnimationBehavior motion = sliderMotionEffect.Behaviors[1];
            float end = timerWidth / slideWidth;
            motion.MotionEffect.Path = "M 0 0 L " + end + " 0 E";
            sliderMotionEffect.Timing.Duration = time;
            sliderMotionEffect.Timing.SmoothStart = Microsoft.Office.Core.MsoTriState.msoFalse;
            sliderMotionEffect.Timing.SmoothEnd = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
    }
}
