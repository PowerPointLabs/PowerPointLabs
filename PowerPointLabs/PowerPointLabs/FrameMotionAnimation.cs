using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{

    class FrameMotionAnimation
    {
        public enum FrameMotionAnimationType { kAutoAnimate, kInSlideAnimate, kStepBackWithBackground, kZoomToAreaPan, kZoomToAreaDeMagnify };
        public static FrameMotionAnimationType animationType = FrameMotionAnimationType.kAutoAnimate;
        public static void AddFrameMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape, float duration)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialRotation = initialShape.Rotation;
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;
            float initialFont = 0.0f;

            float finalX = (finalShape.Left + (finalShape.Width) / 2);
            float finalY = (finalShape.Top + (finalShape.Height) / 2);
            float finalRotation = finalShape.Rotation;
            float finalWidth = finalShape.Width;
            float finalHeight = finalShape.Height;
            float finalFont = 0.0f;

            if (initialShape.HasTextFrame == Office.MsoTriState.msoTrue && (initialShape.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || initialShape.TextFrame.HasText == Office.MsoTriState.msoTrue) && initialShape.TextFrame.TextRange.Font.Size != finalShape.TextFrame.TextRange.Font.Size)
            {
                finalFont = finalShape.TextFrame.TextRange.Font.Size;
                initialFont = initialShape.TextFrame.TextRange.Font.Size;
            }

            int numFrames = (int)(duration / 0.04f);
            numFrames = (numFrames > 30) ? 30 : numFrames;

            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
            float incrementRotation = PowerPointLabsGlobals.GetMinimumRotation(initialRotation, finalRotation) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;
            float incrementFont = (finalFont - initialFont) / numFrames;

            AddFrameAnimationEffects(animationSlide, initialShape, incrementLeft, incrementTop, incrementWidth, incrementHeight, incrementRotation, incrementFont, duration, numFrames);
        }

        public static void AddStepBackFrameMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;

            float finalX = PowerPointPresentation.SlideWidth / 2;
            float finalY = PowerPointPresentation.SlideHeight / 2;
            float finalWidth = PowerPointPresentation.SlideWidth;
            float finalHeight = PowerPointPresentation.SlideHeight;

            int numFrames = 10;
            float duration = numFrames * 0.04f;

            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;

            AddFrameAnimationEffects(animationSlide, initialShape, incrementLeft, incrementTop, incrementWidth, incrementHeight, 0.0f, 0.0f, duration, numFrames);
        }

        public static void AddZoomToAreaPanFrameMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;

            float finalX = (finalShape.Left + (finalShape.Width) / 2);
            float finalY = (finalShape.Top + (finalShape.Height) / 2);
            float finalWidth = finalShape.Width;
            float finalHeight = finalShape.Height;

            int numFrames = 10;
            float duration = numFrames * 0.04f;

            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;

            AddFrameAnimationEffects(animationSlide, initialShape, incrementLeft, incrementTop, incrementWidth, incrementHeight, 0.0f, 0.0f, duration, numFrames);
        }

        private static void AddFrameAnimationEffects(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, float incrementLeft, float incrementTop, float incrementWidth, float incrementHeight, float incrementRotation ,float incrementFont, float duration, int numFrames)
        {
            PowerPoint.Shape lastShape = initialShape;
            PowerPoint.Sequence sequence = animationSlide.TimeLine.MainSequence;
            for (int i = 1; i <= numFrames; i++)
            {
                PowerPoint.Shape dupShape = initialShape.Duplicate()[1];
                if (i != 1 && animationType != FrameMotionAnimationType.kZoomToAreaDeMagnify)
                    sequence[sequence.Count].Delete();

                if (animationType == FrameMotionAnimationType.kInSlideAnimate || animationType == FrameMotionAnimationType.kZoomToAreaPan || animationType == FrameMotionAnimationType.kZoomToAreaDeMagnify)
                    animationSlide.DeleteShapeAnimations(dupShape);

                if (animationType == FrameMotionAnimationType.kZoomToAreaPan)
                    dupShape.Name = "PPTLabsMagnifyPanAreaGroup" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                dupShape.Left = initialShape.Left;
                dupShape.Top = initialShape.Top;

                if (incrementWidth != 0.0f)
                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);

                if (incrementHeight != 0.0f)
                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);

                if (incrementRotation != 0.0f)
                    dupShape.Rotation += (incrementRotation * i);

                if (incrementLeft != 0.0f)
                    dupShape.Left += (incrementLeft * i);

                if (incrementTop != 0.0f)
                    dupShape.Top += (incrementTop * i);

                if (incrementFont != 0.0f)
                    dupShape.TextFrame.TextRange.Font.Size += (incrementFont * i);

                if (i == 1 && (animationType == FrameMotionAnimationType.kInSlideAnimate || animationType == FrameMotionAnimationType.kZoomToAreaPan)) 
                {
                    PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                }
                else
                {
                    PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    appear.Timing.TriggerDelayTime = ((duration / numFrames) * i);
                }

                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                disappear.Exit = Office.MsoTriState.msoTrue;
                disappear.Timing.TriggerDelayTime = ((duration / numFrames) * i);

                lastShape = dupShape;
            }

            if (animationType == FrameMotionAnimationType.kInSlideAnimate || animationType == FrameMotionAnimationType.kZoomToAreaPan || animationType == FrameMotionAnimationType.kZoomToAreaDeMagnify)
            {
                PowerPoint.Effect disappearLast = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                disappearLast.Exit = Office.MsoTriState.msoTrue;
                disappearLast.Timing.TriggerDelayTime = duration;
            }
        }
    }
}
