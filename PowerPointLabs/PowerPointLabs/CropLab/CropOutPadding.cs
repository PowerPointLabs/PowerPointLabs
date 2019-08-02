using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using PowerPointLabs.ActionFramework.Common.Extension;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    internal class CropOutPadding
    {
        private const float Epsilon = 0.05f;
        private static readonly string TempPngFileExportPath = Path.GetTempPath() + @"\cropoutpaddingtemp.png";

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange)
        {
            bool hasChange = false;

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                PowerPoint.Shape shape = shapeRange[i];

                // Store initial properties
                float currentRotation = shape.Rotation;
                float cropLeft = shape.PictureFormat.CropLeft;
                float cropRight = shape.PictureFormat.CropRight;
                float cropTop = shape.PictureFormat.CropTop;
                float cropBottom = shape.PictureFormat.CropBottom;

                // Set properties to zero to do proper calculations
                shape.PictureFormat.CropLeft = 0;
                shape.PictureFormat.CropRight = 0;
                shape.PictureFormat.CropTop = 0;
                shape.PictureFormat.CropBottom = 0;
                shape.Rotation = 0;

                // Get unscaled dimensions
                PowerPoint.ShapeRange origShape = shape.Duplicate();
                origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                float origWidth = origShape.Width;
                float origHeight = origShape.Height;
                origShape.SafeDelete();

                Rectangle origImageRect = new Rectangle();
                Rectangle croppedImageRect = new Rectangle();

                Utils.GraphicsUtil.ExportShape(shape, TempPngFileExportPath);
                using (Bitmap shapeBitmap = new Bitmap(TempPngFileExportPath))
                {
                    origImageRect = new Rectangle(0, 0, shapeBitmap.Width, shapeBitmap.Height);
                    try
                    {
                        croppedImageRect = GetImageBoundingRect(shapeBitmap, shape.Name);
                    }
                    catch (NotSupportedException e)
                    {
                        throw e;
                    }
                }

                float cropRatioLeft = croppedImageRect.Left / (float)origImageRect.Width;
                float cropRatioRight = (origImageRect.Width - croppedImageRect.Width) / (float)origImageRect.Width;
                float cropRatioTop = croppedImageRect.Top / (float)origImageRect.Height;
                float cropRatioBottom = (origImageRect.Height - croppedImageRect.Height) / (float)origImageRect.Height;

                float newCropLeft = Math.Max(origWidth * cropRatioLeft, cropLeft);
                float newCropRight = Math.Max(origWidth * cropRatioRight, cropRight);
                float newCropTop = Math.Max(origHeight * cropRatioTop, cropTop);
                float newCropBottom = Math.Max(origHeight * cropRatioBottom, cropBottom);

                if (!hasChange && 
                    (!IsApproximatelySame(newCropLeft, cropLeft) || 
                    !IsApproximatelySame(newCropRight, cropRight) ||
                    !IsApproximatelySame(newCropTop, cropTop) ||
                    !IsApproximatelySame(newCropBottom, cropBottom)))
                {
                    hasChange = true;
                }
                
                shape.Rotation = currentRotation;
                shape.PictureFormat.CropLeft = newCropLeft;
                shape.PictureFormat.CropRight = newCropRight;
                shape.PictureFormat.CropTop = newCropTop;
                shape.PictureFormat.CropBottom = newCropBottom;
            }

            if (!hasChange)
            {
                throw new CropLabException(CropLabErrorHandler.ErrorCodeNoPaddingCropped.ToString());
            }

            return shapeRange;
        }

        private static bool IsApproximatelySame(float a, float b)
        {
            return Math.Abs(a - b) < Epsilon;
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

        private static Rectangle GetImageBoundingRect(Bitmap bmp, string shapeName)
        {
            if (bmp.PixelFormat != PixelFormat.Format32bppArgb)
            {
                string errorMsg = "Non-Format32bppArgb bitmap for " + shapeName + " is not supported.";
                throw new NotSupportedException(errorMsg);
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
                    right = x + 1;
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
                    bottom = y + 1;
                    break;
                }
            }

            Rectangle boundingRect = new Rectangle(left, top, right, bottom);
            bmp.UnlockBits(bmpData);
            return boundingRect;
        }
    }
}
