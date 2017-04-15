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

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointStepBackSlide(slide);
        }

        public void PrepareForStepBack()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public void AddStepBackAnimationBackground(PowerPoint.Shape shapeToZoom, PowerPoint.Shape backgroundShape, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            DefaultMotionAnimation.AddStepBackMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            DefaultMotionAnimation.AddStepBackMotionAnimation(this, backgroundShape, shapeToZoom, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            
            DefaultMotionAnimation.PreloadShape(this, backgroundShape, false);
            DefaultMotionAnimation.DuplicateAsCoverImage(this, shapeToZoom);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        public void AddStepBackAnimationNonBackground(PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            DefaultMotionAnimation.AddStepBackMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }
    }
}
