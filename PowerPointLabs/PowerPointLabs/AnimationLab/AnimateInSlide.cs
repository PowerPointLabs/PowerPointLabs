using System;
using System.Linq;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.AnimationLab
{
#pragma warning disable 0618
    internal static class AnimateInSlide
    {
        public static void AddAnimationInSlide(bool isHighlightBullets = false, bool isHighlightTextFragments = false)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                currentSlide.RemoveAnimationsForShapes(selectedShapes.Cast<PowerPoint.Shape>().ToList());

                if (!isHighlightBullets && !isHighlightTextFragments)
                {
                    FormatInSlideAnimateShapes(selectedShapes, isHighlightTextFragments);
                }

                if (selectedShapes.Count == 1)
                {
                    InSlideAnimateSingleShape(currentSlide, selectedShapes[1]);
                }
                else
                {
                    InSlideAnimateMultiShape(currentSlide, selectedShapes, isHighlightTextFragments);
                }

                if (!isHighlightBullets && !isHighlightTextFragments)
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                    PowerPointPresentation.Current.AddAckSlide();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddAnimationInSlide");
                throw;
            }
        }

        private static void FormatInSlideAnimateShapes(PowerPoint.ShapeRange shapes, bool isHighlightTextFragments)
        {
            foreach (PowerPoint.Shape sh in shapes) 
            {
                if (isHighlightTextFragments)
                {
                    sh.Name = "PPTLabsHighlightTextFragmentShape" + Guid.NewGuid().ToString();
                }
                else
                {
                    sh.Name = "InSlideAnimateShape" + Guid.NewGuid().ToString();
                }
            }
        }

        private static void InSlideAnimateSingleShape(PowerPointSlide currentSlide, PowerPoint.Shape shapeToAnimate)
        {
            PowerPoint.Effect appear = currentSlide.TimeLine.MainSequence.AddEffect(
                shapeToAnimate, 
                PowerPoint.MsoAnimEffect.msoAnimEffectAppear, 
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            PowerPoint.Effect disappear = currentSlide.TimeLine.MainSequence.AddEffect(
                shapeToAnimate, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, 
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            disappear.Exit = Office.MsoTriState.msoTrue;
        }

        private static void InSlideAnimateMultiShape(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapesToAnimate, 
                                                    bool isHighlightTextFragments)
        {
            for (int num = 1; num <= shapesToAnimate.Count - 1; num++)
            {
                PowerPoint.Shape shape1 = shapesToAnimate[num];
                PowerPoint.Shape shape2 = shapesToAnimate[num + 1];

                if (shape1 == null || shape2 == null)
                {
                    return;
                }

                if (isHighlightTextFragments)
                {
                    //Transition from shape1 to shape2 with movement
                    PowerPoint.Effect shape2Appear = currentSlide.TimeLine.MainSequence.AddEffect(
                        shape2,
                        PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                        PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                }
                else
                {
                    AnimateMovementBetweenShapes(currentSlide, shape1, shape2);

                    //Transition from shape1 to shape2 with fade
                    PowerPoint.Effect shape2Appear = currentSlide.TimeLine.MainSequence.AddEffect(
                        shape2,
                        PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                        PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                }
                PowerPoint.Effect shape1Disappear = currentSlide.TimeLine.MainSequence.AddEffect(
                        shape1.IsStraightLine() ? shape1.ParentGroup : shape1,
                        PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                        PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                shape1Disappear.Exit = Office.MsoTriState.msoTrue;
            }
        }

        private static void AnimateMovementBetweenShapes(PowerPointSlide currentSlide, PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            if (NeedsFrameAnimation(shape1, shape2))
            {
                FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kInSlideAnimate;
                FrameMotionAnimation.AddFrameMotionAnimation(currentSlide, shape1, shape2, AnimationLabSettings.AnimationDuration);
            }
            else
            {
                DefaultMotionAnimation.AddDefaultMotionAnimation(
                    currentSlide,
                    shape1,
                    shape2,
                    AnimationLabSettings.AnimationDuration,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            }
        }

        private static bool NeedsFrameAnimation(PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            float finalFont = 0.0f;
            float initialFont = 0.0f;

            if (shape1.HasTextFrame == Office.MsoTriState.msoTrue && 
                (shape1.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || shape1.TextFrame.HasText == Office.MsoTriState.msoTrue) && 
                shape1.TextFrame.TextRange.Font.Size != shape2.TextFrame.TextRange.Font.Size)
            {
                finalFont = shape2.TextFrame.TextRange.Font.Size;
                initialFont = shape1.TextFrame.TextRange.Font.Size;
            }

            if ((AnimationLabSettings.IsUseFrameAnimation && (shape2.Height != shape1.Height || shape2.Width != shape1.Width)) || 
                ((shape2.Rotation != shape1.Rotation || shape1.Rotation % 90 != 0) && (shape2.Height != shape1.Height || shape2.Width != shape1.Width)) || 
                (!shape1.IsStraightLine() && (shape1.HorizontalFlip != shape2.HorizontalFlip || shape1.VerticalFlip != shape2.VerticalFlip)) || 
                finalFont != initialFont)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
