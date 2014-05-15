using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointStepBackSlide : PowerPointSlide
    {
        private PowerPointStepBackSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsZoomOut" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointStepBackSlide(slide);
        }

        public void PrepareForStepBack()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public void AddStepBackAnimation(PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            if (AutoZoom.backgroundZoomChecked)
            {
                FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kStepBackWithBackground;
                FrameMotionAnimation.AddStepBackFrameMotionAnimation(this, shapeToZoom);
            }
            else
                DefaultMotionAnimation.AddStepBackMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }
    }
}
