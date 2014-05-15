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
        public static void AddDefaultMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape, float duration, PowerPoint.MsoAnimTriggerType trigger)
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

            AddMotionAnimation(animationSlide, initialShape, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, initialShape, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
            AddRotationAnimation(animationSlide, initialShape, initialRotation, finalRotation, duration, ref trigger);
        }

        public static void AddDrillDownMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            float finalWidth = PowerPointPresentation.SlideWidth;
            float initialWidth = referenceShape.Width;
            float finalHeight = PowerPointPresentation.SlideHeight;
            float initialHeight = referenceShape.Height;

            float finalX = (PowerPointPresentation.SlideWidth / 2) * (finalWidth / initialWidth);
            float initialX = (referenceShape.Left + (referenceShape.Width) / 2) * (finalWidth / initialWidth);
            float finalY = (PowerPointPresentation.SlideHeight / 2) * (finalHeight / initialHeight);
            float initialY = (referenceShape.Top + (referenceShape.Height) / 2) * (finalHeight / initialHeight);

            AddMotionAnimation(animationSlide, shapeToZoom, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, shapeToZoom, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
        }

        public static void AddStepBackMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            float initialX = (shapeToZoom.Left + (shapeToZoom.Width) / 2);
            float finalX = (referenceShape.Left + (referenceShape.Width) / 2);
            float initialY = (shapeToZoom.Top + (shapeToZoom.Height) / 2);
            float finalY = (referenceShape.Top + (referenceShape.Height) / 2);

            float initialWidth = shapeToZoom.Width;
            float finalWidth = referenceShape.Width;
            float initialHeight = shapeToZoom.Height;
            float finalHeight = referenceShape.Height;

            AddMotionAnimation(animationSlide, shapeToZoom, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, shapeToZoom, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
        }

        public static void AddZoomToAreaMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape shapeToZoom, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            float initialWidth = initialShape.Width;
            float finalWidth = finalShape.Width;
            float initialHeight = initialShape.Height;
            float finalHeight = finalShape.Height;

            float initialX = (initialShape.Left + (initialShape.Width) / 2) * (finalWidth / initialWidth);
            float finalX = (finalShape.Left + (finalShape.Width) / 2) * (finalWidth / initialWidth);
            float initialY = (initialShape.Top + (initialShape.Height) / 2) * (finalHeight / initialHeight);
            float finalY = (finalShape.Top + (finalShape.Height) / 2) * (finalHeight / initialHeight);

            AddMotionAnimation(animationSlide, shapeToZoom, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, shapeToZoom, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
        }

        private static void AddMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialX, float initialY, float finalX, float finalY, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            if ((finalX != initialX) || (finalY != initialY))
            {
                PowerPoint.Effect effectMotion = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
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
        }

        private static void AddRotationAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialRotation, float finalRotation, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            if (finalRotation != initialRotation)
            {
                PowerPoint.Effect effectRotate = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                effectRotate.Timing.Duration = duration;
                effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
        }

        private static void AddResizeAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialWidth, float initialHeight, float finalWidth, float finalHeight, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
            {
                animationShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                PowerPoint.Effect effectResize = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                effectResize.Timing.Duration = duration;

                resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

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
