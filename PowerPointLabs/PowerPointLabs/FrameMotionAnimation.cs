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
            float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;
            float incrementFont = (finalFont - initialFont) / numFrames;

            PowerPoint.Shape lastShape = initialShape;
            PowerPoint.Sequence sequence = animationSlide.TimeLine.MainSequence;
            for (int i = 1; i <= numFrames; i++)
            {
                PowerPoint.Shape dupShape = initialShape.Duplicate()[1];
                if (i != 1)
                    sequence[sequence.Count].Delete();
  
                PowerPoint.Effect shapeEffect = GetShapeAnimations(animationSlide, dupShape);
                if (shapeEffect != null)
                    shapeEffect.Delete();

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

                if (i == 1)
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

            PowerPoint.Effect disappearLast = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            disappearLast.Exit = Office.MsoTriState.msoTrue;
            disappearLast.Timing.TriggerDelayTime = duration;
        }

        private static float GetMinimumRotation(float fromAngle, float toAngle)
        {
            fromAngle = Normalize(fromAngle);
            toAngle = Normalize(toAngle);

            float rotation1 = toAngle - fromAngle;
            float rotation2 = rotation1 == 0.0f ? 0.0f : Math.Abs(360.0f - Math.Abs(rotation1)) * (rotation1 / Math.Abs(rotation1)) * -1.0f;

            if (Math.Abs(rotation1) < Math.Abs(rotation2))
                return rotation1;
            else
                return rotation2;
        }

        private static float Normalize(float i)
        {
            //find effective angle
            float d = Math.Abs(i) % 360.0f;

            if (i < 0)
                return 360.0f - d; //return positive equivalent
            else
                return d;
        }

        private static PowerPoint.Effect GetShapeAnimations(PowerPointSlide slide, PowerPoint.Shape shape)
        {
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            PowerPoint.Effect e = null;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                {
                    e = effect;
                    break;
                }
            }
            return e;
        }
    }
}
