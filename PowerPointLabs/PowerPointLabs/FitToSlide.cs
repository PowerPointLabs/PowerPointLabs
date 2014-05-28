using System;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class FitToSlide
    {
        private const int TopMost = 0;
        private const int LeftMost = 0;

        public static void DoFitToHeight()
        {
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape);
            float resizeFactor = GetResizeFactorForFitToHeight(selectedShape);

            FitToHeight(selectedShape, resizeFactor, shapeSizeRatio);
            MoveToCenterForFitToHeight(selectedShape);
            AdjustPositionForFitToHeight(selectedShape);
        }

        public static void DoFitToWidth()
        {
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape);
            float resizeFactor = GetResizeFactorForFitToWidth(selectedShape);

            FitToWidth(selectedShape, resizeFactor, shapeSizeRatio);
            MoveToCenterForFitToWidth(selectedShape);
            AdjustPositionForFitToWidth(selectedShape);
        }

        private static void FitToHeight(PowerPoint.Shape shape, float resizeFactor, float shapeSizeRatio)
        {
            shape.Height = PowerPointPresentation.SlideHeight / resizeFactor;
            shape.Width = shape.Height / shapeSizeRatio;
        }

        private static void FitToWidth(PowerPoint.Shape shape, float resizeFactor, float shapeSizeRatio)
        {
            shape.Height = PowerPointPresentation.SlideWidth / resizeFactor;
            shape.Width = shape.Height / shapeSizeRatio;
        }

        private static void MoveToCenterForFitToHeight(PowerPoint.Shape selectedShape)
        {
            selectedShape.Left = (PowerPointPresentation.SlideWidth - selectedShape.Width)/2;
            selectedShape.Top = TopMost;
        }

        private static void AdjustPositionForFitToHeight(PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = GetRotationValueForAdjustPosition(shape);
            float diagonal = GetDiagonal(shape);
            float oppositeSideLength = GetOppositeSideLength(diagonal, rotation);
            float angle1 = (float)Math.Atan(shape.Width / shape.Height);
            float angle2 = (float)((Math.PI - rotation) / 2);

            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Height;
            }
            shape.Top += adjustLength;
        }

        private static void MoveToCenterForFitToWidth(PowerPoint.Shape selectedShape)
        {
            selectedShape.Top = (PowerPointPresentation.SlideHeight - selectedShape.Height)/2;
            selectedShape.Left = LeftMost;
        }

        private static void AdjustPositionForFitToWidth(PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = GetRotationValueForAdjustPosition(shape);
            float diagonal = GetDiagonal(shape);
            float oppositeSideLength = GetOppositeSideLength(diagonal, rotation);
            float angle1 = (float)Math.Atan(shape.Height / shape.Width);
            float angle2 = (float)((Math.PI - rotation) / 2);

            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Width;
            }
            shape.Left += adjustLength;
        }

        private static float GetOppositeSideLength(float diagonal, float rotation)
        {
            //Law of cosines
            return (float)Math.Sqrt((Math.Pow(diagonal, 2) * 2 * (1 - Math.Cos(rotation))));
        }

        private static float GetDiagonal(PowerPoint.Shape shape)
        {
            return (float)Math.Sqrt(Math.Pow(shape.Width / 2, 2) + Math.Pow(shape.Height / 2, 2));
        }

        private static float GetRotationValueForAdjustPosition(PowerPoint.Shape shape)
        {
            float rotation = shape.Rotation;
            if (shape.Rotation > 180 && shape.Rotation <= 360)
            {
                rotation = 360 - shape.Rotation;
            }
            return ConvertDegToRad(rotation);
        }

        private static float GetResizeFactorForFitToWidth(PowerPoint.Shape shape)
        {
            float rotation = GetRotationValueForResizeFactor(shape);
            //calculate resizeFactor for Fit to Height
            float shapeSizeRatio = GetSizeRatio(shape);
            float factor = (float)(Math.Sin(rotation) + Math.Cos(rotation) / shapeSizeRatio);
            return factor;
        }

        private static float GetResizeFactorForFitToHeight(PowerPoint.Shape shape)
        {
            float rotation = GetRotationValueForResizeFactor(shape);
            float shapeSizeRatio = GetSizeRatio(shape);
            float factor = (float)(Math.Cos(rotation) + Math.Sin(rotation) / shapeSizeRatio);
            return factor;
        }

        private static float GetRotationValueForResizeFactor(PowerPoint.Shape shape)
        {
            float rotation;
            if ((int) shape.Rotation == 90)
            {
                rotation = shape.Rotation;
            }
            else if ((int) shape.Rotation == 270)
            {
                rotation = 360 - shape.Rotation;
            }
            else if ((shape.Rotation > 90 && shape.Rotation <= 180)
                     || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                rotation = (360 - shape.Rotation) % 90;
            }
            else
            {
                rotation = shape.Rotation % 90;
            }
            return ConvertDegToRad(rotation);
        }

        private static float ConvertDegToRad(float rotation)
        {
            rotation = (float)((rotation) * Math.PI / 180); return rotation;
        }

        private static float GetSizeRatio(PowerPoint.Shape shape)
        {
            return shape.Height / shape.Width;
        }

        private static PowerPoint.Shape GetSelectedShape()
        {
            return PowerPointPresentation.CurrentSelection.ShapeRange[1];
        }
    }
}
