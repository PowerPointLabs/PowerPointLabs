using System;
using System.Collections.Generic;
using System.Globalization;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class DefaultMotionAnimation
    {
#pragma warning disable 0618
        //Use initial shape and final shape to calculate intial and final positions
        //Add motion, resize and rotation animations to shape
        public static void AddDefaultMotionAnimation(
            PowerPointSlide animationSlide, 
            PowerPoint.Shape initialShape, 
            PowerPoint.Shape finalShape, 
            float duration,
            PowerPoint.MsoAnimTriggerType trigger)
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

            if (Utils.Graphics.IsStraightLine(initialShape))
            {
                double initialAngle = GetLineAngle(initialShape);
                double finalAngle = GetLineAngle(finalShape);
                double deltaAngle = initialAngle - finalAngle;
                finalRotation = (float)(RadiansToDegrees(deltaAngle));

                float initialLength = (float)Math.Sqrt(initialWidth * initialWidth + initialHeight * initialHeight);
                float finalLength = (float)Math.Sqrt(finalWidth * finalWidth + finalHeight * finalHeight);
                float initialAngleCosine = initialWidth / initialLength;
                float initialAngleSine = initialHeight / initialLength;
                finalWidth = finalLength * initialAngleCosine;
                finalHeight = finalLength * initialAngleSine;

                initialShape = MakePivotCenteredLine(animationSlide, initialShape);
            }

            AddMotionAnimation(animationSlide, initialShape, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, initialShape, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
            AddRotationAnimation(animationSlide, initialShape, initialRotation, finalRotation, duration, ref trigger);
        }

        //Use reference shape and slide dimensions to calculate intial and final positions
        //Add motion and resize animations to shapeToZoom
        public static void AddDrillDownMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape shapeToZoom, PowerPoint.Shape referenceShape, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            float finalWidth = PowerPointPresentation.Current.SlideWidth;
            float initialWidth = referenceShape.Width;
            float finalHeight = PowerPointPresentation.Current.SlideHeight;
            float initialHeight = referenceShape.Height;

            float finalX = (PowerPointPresentation.Current.SlideWidth / 2) * (finalWidth / initialWidth);
            float initialX = (referenceShape.Left + (referenceShape.Width) / 2) * (finalWidth / initialWidth);
            float finalY = (PowerPointPresentation.Current.SlideHeight / 2) * (finalHeight / initialHeight);
            float initialY = (referenceShape.Top + (referenceShape.Height) / 2) * (finalHeight / initialHeight);

            AddMotionAnimation(animationSlide, shapeToZoom, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, shapeToZoom, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
        }

        //Use shapeToZoom and reference shape to calculate intial and final positions
        //Add motion and resize animations to shapeToZoom
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

        //Use initial shape and final shape to calculate intial and final positions.
        //Add motion and resize animations to shapeToZoom
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

        /// <summary>
        /// Pans initialShape to the location and size of finalShape
        /// </summary>
        public static void AddZoomToAreaPanAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.Shape finalShape, PowerPoint.MsoAnimTriggerType trigger)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;

            float finalX = (finalShape.Left + (finalShape.Width) / 2);
            float finalY = (finalShape.Top + (finalShape.Height) / 2);
            float finalWidth = finalShape.Width;
            float finalHeight = finalShape.Height;

            float duration = 0.4f;

            AddMotionAnimation(animationSlide, initialShape, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, initialShape, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);
        }

        /// <summary>
        /// Zoom out from initial shape to the full slide.
        /// Returns the final disappear effect for initialShape.
        /// </summary>
        public static PowerPoint.Effect AddZoomOutMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape initialShape, PowerPoint.MsoAnimTriggerType trigger)
        {
            float initialX = (initialShape.Left + (initialShape.Width) / 2);
            float initialY = (initialShape.Top + (initialShape.Height) / 2);
            float initialWidth = initialShape.Width;
            float initialHeight = initialShape.Height;

            float finalX = PowerPointPresentation.Current.SlideWidth / 2;
            float finalY = PowerPointPresentation.Current.SlideHeight / 2;
            float finalWidth = PowerPointPresentation.Current.SlideWidth;
            float finalHeight = PowerPointPresentation.Current.SlideHeight;

            float duration = 0.4f;

            AddMotionAnimation(animationSlide, initialShape, initialX, initialY, finalX, finalY, duration, ref trigger);
            AddResizeAnimation(animationSlide, initialShape, initialWidth, initialHeight, finalWidth, finalHeight, duration, ref trigger);

            var sequence = animationSlide.TimeLine.MainSequence;
            var effectDisappear = sequence.AddEffect(initialShape, PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0.01f;
            return effectDisappear;
        }

        /// <summary>
        /// Preloads a shape within the slide to reduce lag. Call after the animations for the shape have been created.
        /// </summary>
        public static void PreloadShape(PowerPointSlide animationSlide, PowerPoint.Shape shape, bool addCoverImage=true)
        {
            // The cover image is used to cover the screen while the preloading happens behind the cover image.
            PowerPoint.Shape coverImage = null;
            if (addCoverImage)
            {
                coverImage = shape.Duplicate()[1];
                coverImage.Left = shape.Left;
                coverImage.Top = shape.Top;
                animationSlide.RemoveAnimationsForShape(coverImage);
            }

            float originalWidth = shape.Width;
            float originalHeight = shape.Height;
            float originalLeft = shape.Left;
            float originalTop = shape.Top;

            // fit the shape exactly in the screen for preloading.
            float scaleRatio = Math.Min(PowerPointPresentation.Current.SlideWidth / shape.Width,
                PowerPointPresentation.Current.SlideHeight / shape.Height);
            animationSlide.RelocateShapeWithoutPath(shape, 0, 0, shape.Width * scaleRatio, shape.Height * scaleRatio);

            var trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            var effectMotion = AddMotionAnimation(animationSlide, shape, shape.Left, shape.Top, originalLeft + (originalWidth - shape.Width) / 2, originalTop + (originalHeight - shape.Height) / 2, 0, ref trigger);
            var effectResize = AddResizeAnimation(animationSlide, shape, shape.Width, shape.Height, originalWidth, originalHeight, 0, ref trigger);

            // Make "cover" image disappear after preload.
            PowerPoint.Effect effectDisappear = null;
            if (addCoverImage)
            {
                var sequence = animationSlide.TimeLine.MainSequence;
                effectDisappear = sequence.AddEffect(coverImage, PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectDisappear.Exit = Office.MsoTriState.msoTrue;
                effectDisappear.Timing.Duration = 0.01f;
            }


            int firstEffectIndex = animationSlide.IndexOfFirstEffect(shape);
            // Move the animations to just before the index of the first effect.
            if (effectDisappear != null) effectDisappear.MoveTo(firstEffectIndex);
            if (effectResize != null) effectResize.MoveTo(firstEffectIndex);
            if (effectMotion != null) effectMotion.MoveTo(firstEffectIndex);
        }

        /// <summary>
        /// Creates a cover image from a copy of the shape to obstruct the viewer while preloading images in the background.
        /// </summary>
        public static void DuplicateAsCoverImage(PowerPointSlide animationSlide, PowerPoint.Shape shape)
        {
            var coverImage = shape.Duplicate()[1];
            coverImage.Left = shape.Left;
            coverImage.Top = shape.Top;
            animationSlide.RemoveAnimationsForShape(coverImage);

            var sequence = animationSlide.TimeLine.MainSequence;
            var effectDisappear = sequence.AddEffect(coverImage, PowerPoint.MsoAnimEffect.msoAnimEffectFade,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0.01f;
            effectDisappear.MoveTo(1);
        }

        private static PowerPoint.Effect AddMotionAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialX, float initialY, float finalX, float finalY, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            if ((finalX != initialX) || (finalY != initialY))
            {
                PowerPoint.Effect effectMotion = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                effectMotion.Timing.Duration = duration;
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;

                //Create VML path for the motion path
                //This path needs to be a curved path to allow the user to edit points
                float point1X = ((finalX - initialX) / 2f) / PowerPointPresentation.Current.SlideWidth;
                float point1Y = ((finalY - initialY) / 2f) / PowerPointPresentation.Current.SlideHeight;
                float point2X = (finalX - initialX) / PowerPointPresentation.Current.SlideWidth;
                float point2Y = (finalY - initialY) / PowerPointPresentation.Current.SlideHeight;
                motion.MotionEffect.Path = "M 0 0 C "
                    + point1X.ToString(CultureInfo.InvariantCulture) + " "
                    + point1Y.ToString(CultureInfo.InvariantCulture) + " "
                    + point1X.ToString(CultureInfo.InvariantCulture) + " "
                    + point1Y.ToString(CultureInfo.InvariantCulture) + " "
                    + point2X.ToString(CultureInfo.InvariantCulture) + " "
                    + point2Y.ToString(CultureInfo.InvariantCulture) + " E";
                effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;
                return effectMotion;
            }
            return null;
        }

        private static PowerPoint.Effect AddRotationAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialRotation, float finalRotation, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            if (finalRotation != initialRotation)
            {
                PowerPoint.Effect effectRotate = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                effectRotate.Timing.Duration = duration;
                effectRotate.EffectParameters.Amount = LegacyShapeUtil.GetMinimumRotation(initialRotation, finalRotation);
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                return effectRotate;
            }
            return null;
        }

        private static PowerPoint.Effect AddResizeAnimation(PowerPointSlide animationSlide, PowerPoint.Shape animationShape, float initialWidth, float initialHeight, float finalWidth, float finalHeight, float duration, ref PowerPoint.MsoAnimTriggerType trigger)
        {
            // To prevent zero multiplication and zero division
            initialWidth = SetToPositiveMinIfIsZero(initialWidth);
            initialHeight = SetToPositiveMinIfIsZero(initialHeight);
            finalWidth = SetToPositiveMinIfIsZero(finalWidth);
            finalHeight = SetToPositiveMinIfIsZero(finalHeight);

            if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
            {
                animationShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                PowerPoint.Effect effectResize = animationSlide.TimeLine.MainSequence.AddEffect(animationShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];

                effectResize.Timing.Duration = duration;

                resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                return effectResize;
            }
            return null;
        }

        private static PowerPoint.Shape MakePivotCenteredLine(PowerPointSlide animationSlide, PowerPoint.Shape line)
        {
            // Add a 180 degree rotated line to offset its pivot shift caused by arrowhead
            PowerPoint.Shape transparentLine = line.Duplicate()[1];
            transparentLine.Line.Transparency = 1.0f;
            transparentLine.Rotation = 180.0f;
            transparentLine.Left = line.Left;
            transparentLine.Top = line.Top;

            PowerPoint.Shape tempAnimationHolder = line.Duplicate()[1];
            animationSlide.TransferAnimation(line, tempAnimationHolder);

            List<PowerPoint.Shape> toGroup = new List<PowerPoint.Shape> { line, transparentLine };
            PowerPoint.Shape newLine = animationSlide.ToShapeRange(toGroup).Group();

            animationSlide.TransferAnimation(tempAnimationHolder, newLine);
            tempAnimationHolder.Delete();

            return newLine;
        }

        private static double GetLineAngle(PowerPoint.Shape line)
        {
            double angle = 0.0;

            if (line.Width == 0.0f)
            {
                angle = Math.PI / 2.0;
            }
            else if (line.Height == 0.0f)
            {
                angle = 0.0;
            }
            else
            {
                angle = Math.Atan(line.Height / line.Width);
            }

            if (line.HorizontalFlip == Office.MsoTriState.msoTrue &&
                line.VerticalFlip == Office.MsoTriState.msoTrue)
            {
                // Pointing top left (2nd quadrant)
                angle = Math.PI - angle;
            }
            else if (line.HorizontalFlip == Office.MsoTriState.msoTrue &&
                     line.VerticalFlip == Office.MsoTriState.msoFalse)
            {
                // Pointing bottom left (3rd quadrant)
                angle = Math.PI + angle;
            }
            else if (line.HorizontalFlip == Office.MsoTriState.msoFalse &&
                     line.VerticalFlip == Office.MsoTriState.msoFalse)
            {
                // Pointing bottom right (4th quadrant)
                angle = Math.PI * 2.0 - angle;
            }

            return angle;
        }

        private static double RadiansToDegrees(double radians)
        {
            return radians * (180.0 / Math.PI);
        }

        private static float SetToPositiveMinIfIsZero(float value)
        {
            return value == 0.0f ? 0.1f : value;
        }
    }
}
