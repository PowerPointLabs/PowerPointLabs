using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointDrillDownSlide : PowerPointSlide
    {
        private PowerPointDrillDownSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsZoomIn" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointDrillDownSlide(slide);
        }

        public void PrepareForDrillDown()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public void AddDrillDownAnimation(PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            if (AutoZoom.backgroundZoomChecked)
                DefaultMotionAnimation.AddDrillDownMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            else
            {
                ManageNonMatchingShapes(shapeToZoom, indicatorShape);
                DefaultMotionAnimation.AddDefaultMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            }
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }

        private void ManageNonMatchingShapes(PowerPoint.Shape shapeToZoom, PowerPoint.Shape indicatorShape)
        {
            foreach (PowerPoint.Shape sh in _slide.Shapes)
            {
                if (!sh.Equals(indicatorShape) && !sh.Equals(shapeToZoom))
                {
                    if (!HasExitAnimation(sh))
                    {
                        DeleteShapeAnimations(sh);
                        PowerPoint.Effect effectFade = _slide.TimeLine.MainSequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        effectFade.Exit = Office.MsoTriState.msoTrue;
                        effectFade.Timing.Duration = AutoAnimate.defaultDuration;
                        //fadeFlag = true;
                    }
                    else
                    {
                        DeleteShapeAnimations(sh);
                        PowerPoint.Effect effectDisappear = null;
                        effectDisappear = _slide.TimeLine.MainSequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        effectDisappear.Exit = Office.MsoTriState.msoTrue;
                        effectDisappear.Timing.Duration = 0;
                    }
                }
            }
        }
    }
}
