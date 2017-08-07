using System;

using PowerPointLabs.AnimationLab;

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

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointDrillDownSlide(slide);
        }

        public void PrepareForDrillDown()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public void AddDrillDownAnimationNoBackground(PowerPoint.Shape backgroundShape, PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            ManageNonMatchingShapes(shapeToZoom, indicatorShape);
            DefaultMotionAnimation.AddDefaultMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            
            DefaultMotionAnimation.PreloadShape(this, shapeToZoom, false);
            DefaultMotionAnimation.DuplicateAsCoverImage(this, backgroundShape);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        public void AddDrillDownAnimationBackground(PowerPoint.Shape backgroundShape, PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            DefaultMotionAnimation.AddDrillDownMotionAnimation(this, backgroundShape, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            DefaultMotionAnimation.AddDefaultMotionAnimation(this, shapeToZoom, backgroundShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            DefaultMotionAnimation.PreloadShape(this, shapeToZoom, false);
            DefaultMotionAnimation.DuplicateAsCoverImage(this, backgroundShape);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
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
                        effectFade.Timing.Duration = AnimationLabSettings.AnimationDuration;
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
