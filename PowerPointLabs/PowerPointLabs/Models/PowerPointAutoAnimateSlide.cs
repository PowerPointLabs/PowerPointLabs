using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointAutoAnimateSlide : PowerPointSlide
    {
        private PowerPointAutoAnimateSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPSlideAnimated" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointAutoAnimateSlide(slide);
        }

        public void PrepareForAutoAnimate()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public void AddAutoAnimation(PowerPoint.Shape[] currentSlideShapes, PowerPoint.Shape[] nextSlideSlideShapes, int[] matchingShapeIDs)
        {
            PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
            ManageNonMatchingShapes(matchingShapeIDs, indicatorShape.Id);
            AnimateMatchingShapes(currentSlideShapes, nextSlideSlideShapes, matchingShapeIDs);
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void AnimateMatchingShapes(PowerPoint.Shape[] currentSlideShapes, PowerPoint.Shape[] nextSlideSlideShapes, int[] matchingShapeIDs)
        {
            int matchingShapeIndex;

            // Copy the shapes as the list may be modified when iterating
            PowerPoint.Shape[] slideShapesCopy = new PowerPoint.Shape[_slide.Shapes.Count];
            for (int i = 0; i < slideShapesCopy.Length; i++)
            {
                slideShapesCopy[i] = _slide.Shapes[i + 1];
            }

            foreach (PowerPoint.Shape sh in slideShapesCopy)
            {
                if (matchingShapeIDs.Contains(sh.Id))
                {
                    matchingShapeIndex = Array.IndexOf(matchingShapeIDs, sh.Id);
                    if (matchingShapeIndex < matchingShapeIDs.Count() && sh.Id == matchingShapeIDs[matchingShapeIndex])
                    {
                        DeleteShapeAnimations(sh);
                        PowerPoint.MsoAnimTriggerType trigger = (matchingShapeIndex == 0) ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious : PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        if (NeedsFrameAnimation(sh, nextSlideSlideShapes[matchingShapeIndex]))
                        {
                            FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kAutoAnimate;
                            FrameMotionAnimation.AddFrameMotionAnimation(this, sh, nextSlideSlideShapes[matchingShapeIndex], AutoAnimate.defaultDuration);
                        }
                        else
                            DefaultMotionAnimation.AddDefaultMotionAnimation(this, sh, nextSlideSlideShapes[matchingShapeIndex], AutoAnimate.defaultDuration, trigger);
                    }
                }
            }
        }

        //Fade out non-matching shapes. If shape has exit animation, then delete it
        private void ManageNonMatchingShapes(int[] matchingShapeIDs, int indicatorShapeID)
        {
            foreach (PowerPoint.Shape sh in _slide.Shapes)
            {
                if (!matchingShapeIDs.Contains(sh.Id) && sh.Id != indicatorShapeID)
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

        private static bool NeedsFrameAnimation(PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            float finalFont = 0.0f;
            float initialFont = 0.0f;

            if (shape1.HasTextFrame == Office.MsoTriState.msoTrue && (shape1.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || shape1.TextFrame.HasText == Office.MsoTriState.msoTrue) && shape1.TextFrame.TextRange.Font.Size != shape2.TextFrame.TextRange.Font.Size)
            {
                finalFont = shape2.TextFrame.TextRange.Font.Size;
                initialFont = shape1.TextFrame.TextRange.Font.Size;
            }

            if ((AutoAnimate.frameAnimationChecked && (shape2.Height != shape1.Height || shape2.Width != shape1.Width))
                || ((shape2.Rotation != shape1.Rotation || shape1.Rotation % 90 != 0) && (shape2.Height != shape1.Height || shape2.Width != shape1.Width))
                || (!Utils.Graphics.IsStraightLine(shape1) && (shape1.HorizontalFlip != shape2.HorizontalFlip || shape1.VerticalFlip != shape2.VerticalFlip))
                || finalFont != initialFont)
            {
                return true;
            }
            else
            {
                return false;
            }
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
