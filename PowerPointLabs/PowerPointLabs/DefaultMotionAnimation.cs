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
    class DefaultMotionAnimation
    {
        public static void AddDefaultMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape, float duration)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialRotation = initialShape.Rotation;
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;

            float finalX = (finalShape.Left + (finalShape.Width) / 2);
            float finalY = (finalShape.Top + (finalShape.Height) / 2);
            float finalRotation = finalShape.Rotation;
            float finalWidth = finalShape.Width;
            float finalHeight = finalShape.Height;

            PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            PowerPoint.Sequence sequence = animationSlide.TimeLine.MainSequence;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectRotate = null;

            if ((finalX != initialX) || (finalY != initialY))
            {
                effectMotion = sequence.AddEffect(initialShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                effectMotion.Timing.Duration = duration;
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;

                //Create VML path for the motion path
                //This path needs to be a curved path to allow the user to edit points
                float point1X = ((finalX - initialX) / 2f) / PowerPointPresentation.SlideWidth;
                float point1Y = ((finalY - initialY) / 2f) / PowerPointPresentation.SlideHeight;
                float point2X = (finalX - initialX) / PowerPointPresentation.SlideWidth;
                float point2Y = (finalY - initialY) / PowerPointPresentation.SlideHeight;
                motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " E";
                effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;
            }

            //Resize Effect
            if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
            {
                initialShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                effectResize = sequence.AddEffect(initialShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                effectResize.Timing.Duration = duration;

                resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }

            //Rotation Effect
            if (finalRotation != initialRotation)
            {
                effectRotate = sequence.AddEffect(initialShape, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                effectRotate.Timing.Duration = duration;
                effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
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
    }
}
