using System;
using PowerPointLabs.ActionFramework.Common.Extension;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    internal class CropToAspectRatio
    {
        private const float Epsilon = 0.05f;

        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, float aspectRatio)
        {
            PowerPoint.ShapeRange croppedShape = Crop(selection.ShapeRange, aspectRatio);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, float aspectRatio)
        {
            bool hasChange = false;

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                float origWidth = origShape.Width;
                float origHeight = origShape.Height;
                origShape.SafeDelete();

                float currentWidth = shapeRange[i].Width - (shapeRange[i].PictureFormat.CropLeft + shapeRange[i].PictureFormat.CropRight) / origWidth;
                float currentHeight = shapeRange[i].Height - (shapeRange[i].PictureFormat.CropTop + shapeRange[i].PictureFormat.CropBottom) / origHeight;
                float currentProportions = currentWidth / currentHeight;

                if (IsApproximatelyEquals(currentProportions, aspectRatio))
                {
                    continue;
                }
                else if (currentProportions > aspectRatio)
                {
                    // Crop the width
                    float desiredWidth = currentHeight * aspectRatio;
                    float widthToCrop = origWidth * ((currentWidth - desiredWidth) / currentWidth);
                    CropHorizontal(shapeRange[i], widthToCrop);
                    hasChange = true;
                }
                else if (currentProportions < aspectRatio)
                {
                    // Crop the height
                    float desiredHeight = currentWidth / aspectRatio;
                    float heightToCrop = origHeight * ((currentHeight - desiredHeight) / currentHeight);
                    CropVertical(shapeRange[i], heightToCrop);
                    hasChange = true;
                }
            }

            if (!hasChange)
            {
                throw new CropLabException(CropLabErrorHandler.ErrorCodeNoAspectRatioCropped.ToString());
            }

            return shapeRange;
        }

        private static void CropHorizontal(PowerPoint.Shape shape, float cropAmount)
        {
            switch (CropLabSettings.AnchorPosition)
            {
                case AnchorPosition.TopLeft:
                case AnchorPosition.MiddleLeft:
                case AnchorPosition.BottomLeft:
                    shape.PictureFormat.CropRight += cropAmount;
                    break;
                case AnchorPosition.Top:
                case AnchorPosition.Middle:
                case AnchorPosition.Bottom:
                case AnchorPosition.Reference:
                    shape.PictureFormat.CropLeft += cropAmount / 2.0f;
                    shape.PictureFormat.CropRight += cropAmount / 2.0f;
                    break;
                case AnchorPosition.TopRight:
                case AnchorPosition.MiddleRight:
                case AnchorPosition.BottomRight:
                    shape.PictureFormat.CropLeft += cropAmount;
                    break;
            }
        }

        private static void CropVertical(PowerPoint.Shape shape, float cropAmount)
        {
            switch (CropLabSettings.AnchorPosition)
            {
                case AnchorPosition.TopLeft:
                case AnchorPosition.Top:
                case AnchorPosition.TopRight:
                    shape.PictureFormat.CropBottom += cropAmount;
                    break;
                case AnchorPosition.MiddleLeft:
                case AnchorPosition.Middle:
                case AnchorPosition.MiddleRight:
                case AnchorPosition.Reference:
                    shape.PictureFormat.CropTop += cropAmount / 2.0f;
                    shape.PictureFormat.CropBottom += cropAmount / 2.0f;
                    break;
                case AnchorPosition.BottomLeft:
                case AnchorPosition.Bottom:
                case AnchorPosition.BottomRight:
                    shape.PictureFormat.CropTop += cropAmount;
                    break;
            }
        }

        private static bool IsApproximatelyEquals(float a, float b)
        {
            return Math.Abs(a - b) < Epsilon;
        }
    }
}
