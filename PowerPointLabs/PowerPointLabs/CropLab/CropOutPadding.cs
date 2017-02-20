using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    internal class CropOutPadding
    {
        private static readonly string TempPngFileExportPath = Path.GetTempPath() + @"\cropoutpaddingtemp.png";

        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, CropLabErrorHandler errorHandler = null)
        {
            if (!VerifyIsSelectionValid(selection, errorHandler))
            {
                return null;
            }

            var croppedShape = Crop(selection.ShapeRange, errorHandler);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, CropLabErrorHandler errorHandler = null)
        {
            if (!VerifyIsShapeRangeValid(shapeRange, errorHandler))
            {
                return null;
            }

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                float currentRotation = shapeRange[i].Rotation;
                float cropLeft = shapeRange[i].PictureFormat.CropLeft;
                float cropRight = shapeRange[i].PictureFormat.CropRight;
                float cropTop = shapeRange[i].PictureFormat.CropTop;
                float cropBottom = shapeRange[i].PictureFormat.CropBottom;

                shapeRange[i].PictureFormat.CropLeft = 0;
                shapeRange[i].PictureFormat.CropRight = 0;
                shapeRange[i].PictureFormat.CropTop = 0;
                shapeRange[i].PictureFormat.CropBottom = 0;
                shapeRange[i].Rotation = 0;

                Utils.Graphics.ExportShape(shapeRange[i], TempPngFileExportPath);
                using (Bitmap shapeImage = new Bitmap(TempPngFileExportPath))
                {
                    PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                    origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                    origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                    float origWidth = origShape.Width;
                    float origHeight = origShape.Height;
                    origShape.Delete();

                    Rectangle cropRect = GetImageBoundingRect(shapeImage);
                    float cropRatioLeft = cropRect.Left / (float)shapeImage.Width;
                    float cropRatioRight = (shapeImage.Width - cropRect.Width) / (float)shapeImage.Width;
                    float cropRatioTop = cropRect.Top / (float)shapeImage.Height;
                    float cropRatioBottom = (shapeImage.Height - cropRect.Height) / (float)shapeImage.Height;

                    float newCropLeft = origWidth * cropRatioLeft;
                    float newCropRight = origWidth * cropRatioRight;
                    float newCropTop = origHeight * cropRatioTop;
                    float newCropBottom = origHeight * cropRatioBottom;

                    cropLeft = cropLeft < newCropLeft ? newCropLeft : cropLeft;
                    cropRight = cropRight < newCropRight ? newCropRight : cropRight;
                    cropTop = cropTop < newCropTop ? newCropTop : cropTop;
                    cropBottom = cropBottom < newCropBottom ? newCropBottom : cropBottom;
                }

                shapeRange[i].Rotation = currentRotation;
                shapeRange[i].PictureFormat.CropLeft = cropLeft;
                shapeRange[i].PictureFormat.CropRight = cropRight;
                shapeRange[i].PictureFormat.CropTop = cropTop;
                shapeRange[i].PictureFormat.CropBottom = cropBottom;
            }

            return shapeRange;
        }

        private static bool IsImageRowTransparent(Bitmap image, int y)
        {
            for (int x = 0; x < image.Width; x++)
            {
                if (image.GetPixel(x, y).A > 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsImageColumnTransparent(Bitmap image, int x)
        {
            for (int y = 0; y < image.Height; y++)
            {
                if (image.GetPixel(x, y).A > 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static Rectangle GetImageBoundingRect(Bitmap image)
        {
            int left = 0;
            int top = 0;
            int right = 0;
            int bottom = 0;

            // Get left boundary
            for (int x = 0; x < image.Width; x++)
            {
                if (!IsImageColumnTransparent(image, x))
                {
                    left = x;
                    break;
                }
            }

            // Get right boundary
            for (int x = image.Width - 1; x >= 0; x--)
            {
                if (!IsImageColumnTransparent(image, x))
                {
                    right = x;
                    break;
                }
            }

            // Get top boundary
            for (int y = 0; y < image.Height; y++)
            {
                if (!IsImageRowTransparent(image, y))
                {
                    top = y;
                    break;
                }
            }

            // Get bottom boundary
            for (int y = image.Height - 1; y >= 0; y--)
            {
                if (!IsImageRowTransparent(image, y))
                {
                    bottom = y;
                    break;
                }
            }

            Rectangle boundingRect = new Rectangle(left, top, right, bottom);
            return boundingRect;
        }

        private static bool VerifyIsSelectionValid(PowerPoint.Selection selection, CropLabErrorHandler errorHandler)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            return true;
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, CropLabErrorHandler errorHandler)
        {
            if (shapeRange.Count < 1)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, errorHandler);
                return false;
            }

            return true;
        }

        private static bool IsPictureForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsPicture);
        }

        private static bool IsPicture(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPicture;
        }

        private static void HandleErrorCode(int errorCode, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }

            switch (errorCode)
            {
                case CropLabErrorHandler.ErrorCodeSelectionIsInvalid:
                    errorHandler.ProcessErrorCode(errorCode, "Crop Out Padding", "1", "picture");
                    break;
                case CropLabErrorHandler.ErrorCodeSelectionMustBePicture:
                    errorHandler.ProcessErrorCode(errorCode, "Crop Out Padding");
                    break;
                default:
                    errorHandler.ProcessErrorCode(errorCode);
                    break;
            }
        }
    }
}
