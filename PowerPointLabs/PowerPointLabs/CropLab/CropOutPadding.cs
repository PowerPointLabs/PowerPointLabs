using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

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
                // Store initial properties
                float currentRotation = shapeRange[i].Rotation;
                float cropLeft = shapeRange[i].PictureFormat.CropLeft;
                float cropRight = shapeRange[i].PictureFormat.CropRight;
                float cropTop = shapeRange[i].PictureFormat.CropTop;
                float cropBottom = shapeRange[i].PictureFormat.CropBottom;

                // Set properties to zero to do proper calculations
                shapeRange[i].PictureFormat.CropLeft = 0;
                shapeRange[i].PictureFormat.CropRight = 0;
                shapeRange[i].PictureFormat.CropTop = 0;
                shapeRange[i].PictureFormat.CropBottom = 0;
                shapeRange[i].Rotation = 0;

                // Get unscaled dimensions
                PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                float origWidth = origShape.Width;
                float origHeight = origShape.Height;
                origShape.Delete();

                Rectangle origImageRect = new Rectangle();
                Rectangle croppedImageRect = new Rectangle();

                Utils.Graphics.ExportShape(shapeRange[i], TempPngFileExportPath);
                using (Bitmap shapeBitmap = new Bitmap(TempPngFileExportPath))
                {
                    origImageRect = new Rectangle(0, 0, shapeBitmap.Width, shapeBitmap.Height);
                    try
                    {
                        croppedImageRect = GetImageBoundingRect(shapeBitmap);
                    }
                    catch (NotSupportedException e)
                    {
                        if (errorHandler != null)
                        {
                            string errorMsg = "An unexpected error occurred in Crop Out Padding for " + 
                                              shapeRange[i].Name + ". " +
                                              "Exported bitmap data should be in Format32bppArgb PNG format.";
                            errorHandler.ProcessException(e, errorMsg);
                        }
                        return null;
                    }    
                }

                float cropRatioLeft = croppedImageRect.Left / (float)origImageRect.Width;
                float cropRatioRight = (origImageRect.Width - croppedImageRect.Width) / (float)origImageRect.Width;
                float cropRatioTop = croppedImageRect.Top / (float)origImageRect.Height;
                float cropRatioBottom = (origImageRect.Height - croppedImageRect.Height) / (float)origImageRect.Height;

                float newCropLeft = origWidth * cropRatioLeft;
                float newCropRight = origWidth * cropRatioRight;
                float newCropTop = origHeight * cropRatioTop;
                float newCropBottom = origHeight * cropRatioBottom;

                // Crop if it is more than current crop
                cropLeft = cropLeft < newCropLeft ? newCropLeft : cropLeft;
                cropRight = cropRight < newCropRight ? newCropRight : cropRight;
                cropTop = cropTop < newCropTop ? newCropTop : cropTop;
                cropBottom = cropBottom < newCropBottom ? newCropBottom : cropBottom;

                // Restore original properties
                shapeRange[i].Rotation = currentRotation;
                shapeRange[i].PictureFormat.CropLeft = cropLeft;
                shapeRange[i].PictureFormat.CropRight = cropRight;
                shapeRange[i].PictureFormat.CropTop = cropTop;
                shapeRange[i].PictureFormat.CropBottom = cropBottom;
            }

            return shapeRange;
        }

        private static bool IsImageRowTransparent(BitmapData bmpData, byte[] argbBuffer, int y)
        {
            for (int x = 0; x < bmpData.Width; x++)
            {
                byte alpha = argbBuffer[y * bmpData.Stride + 4 * x + 3];
                if (alpha != 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsImageColumnTransparent(BitmapData bmpData, byte[] argbBuffer, int x)
        {
            for (int y = 0; y < bmpData.Height; y++)
            {
                byte alpha = argbBuffer[y * bmpData.Stride + 4 * x + 3];
                if (alpha != 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static Rectangle GetImageBoundingRect(Bitmap bmp)
        {
            if (bmp.PixelFormat != PixelFormat.Format32bppArgb)
            {
                throw new NotSupportedException("Non-Format32bppArgb bitmaps are not supported.");
            }
            
            int left = -1;
            int top = -1;
            int right = -1;
            int bottom = -1;

            // Lock the bitmap data into system memory for faster read/write
            BitmapData bmpData = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);
            int bytesCount = Math.Abs(bmpData.Stride) * bmp.Height;
            byte[] argbBuffer = new byte[bytesCount];
            System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, argbBuffer, 0, bytesCount);

            // Get left boundary
            for (int x = 0; x < bmpData.Width; x++)
            {
                if (!IsImageColumnTransparent(bmpData, argbBuffer, x))
                {
                    left = x;
                    break;
                }
            }

            // Return immediately if entire image is transparent
            if (left == -1)
            {
                bmp.UnlockBits(bmpData);
                return new Rectangle(0, 0, 0, 0);
            }

            // Get right boundary
            for (int x = bmpData.Width - 1; x >= left; x--)
            {
                if (!IsImageColumnTransparent(bmpData, argbBuffer, x))
                {
                    right = x;
                    break;
                }
            }

            // Get top boundary
            for (int y = 0; y < bmpData.Height; y++)
            {
                if (!IsImageRowTransparent(bmpData, argbBuffer, y))
                {
                    top = y;
                    break;
                }
            }

            // Get bottom boundary
            for (int y = bmpData.Height - 1; y >= top; y--)
            {
                if (!IsImageRowTransparent(bmpData, argbBuffer, y))
                {
                    bottom = y;
                    break;
                }
            }

            Rectangle boundingRect = new Rectangle(left, top, right, bottom);
            bmp.UnlockBits(bmpData);
            return boundingRect;
        }

        private static bool VerifyIsSelectionValid(PowerPoint.Selection selection, CropLabErrorHandler errorHandler)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            return true;
        }

        private static bool VerifyIsShapeRangeValid(PowerPoint.ShapeRange shapeRange, CropLabErrorHandler errorHandler)
        {
            if (shapeRange.Count < 1)
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, errorHandler);
                return false;
            }

            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, errorHandler);
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

        private static void HandleErrorCodeIfRequired(int errorCode, CropLabErrorHandler errorHandler)
        {
            if (errorHandler == null)
            {
                return;
            }
            errorHandler.ProcessErrorCode(errorCode, "Crop Out Padding", CropLabErrorHandler.SelectionTypePicture, 1);
        }
    }
}
