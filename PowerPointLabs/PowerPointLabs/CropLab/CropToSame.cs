using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Utils;

namespace PowerPointLabs.CropLab
{
    public class CropToSame
    {

        private const float Epsilon = 0.00001F; // Prevents divide by zero

        public static bool CropSelection(ShapeRange shapeRange)
        {
            bool hasChange = false;

            Shape refObj = shapeRange[1];

            float refScaleWidth = refObj.GetScaleWidth();
            float refScaleHeight = refObj.GetScaleHeight();

            float cropTop = Epsilon;
            float cropBottom = Epsilon;
            float cropLeft = Epsilon;
            float cropRight = Epsilon;

            if (!refObj.IsShape())
            {
                cropTop = Math.Max(refObj.PictureFormat.CropTop, Epsilon);
                cropBottom = Math.Max(refObj.PictureFormat.CropBottom, Epsilon);
                cropLeft = Math.Max(refObj.PictureFormat.CropLeft, Epsilon);
                cropRight = Math.Max(refObj.PictureFormat.CropRight, Epsilon);
            }

            float refShapeCroppedHeight = cropTop + cropBottom;
            float refShapeCroppedWidth = cropLeft + cropRight;

            for (int i = 2; i <= shapeRange.Count; i++)
            {
                float heightToCrop = shapeRange[i].Height - refObj.Height;
                float widthToCrop = shapeRange[i].Width - refObj.Width;
                if (heightToCrop <= 0 && widthToCrop <= 0)
                {
                    continue;
                }
                hasChange = true;

                float scaleWidth = shapeRange[i].GetScaleWidth();
                float scaleHeight = shapeRange[i].GetScaleHeight();
                if (CropLabSettings.AnchorPosition == AnchorPosition.Reference)
                {
                    shapeRange[i].PictureFormat.CropTop += Math.Max(0, heightToCrop * cropTop / refShapeCroppedHeight / scaleHeight);
                    shapeRange[i].PictureFormat.CropLeft += Math.Max(0, widthToCrop * cropLeft / refShapeCroppedWidth / scaleWidth);
                    shapeRange[i].PictureFormat.CropRight += Math.Max(0, widthToCrop * cropRight / refShapeCroppedWidth / scaleWidth);
                    shapeRange[i].PictureFormat.CropBottom += Math.Max(0, heightToCrop * cropBottom / refShapeCroppedHeight / scaleHeight);
                }
                else
                {
                    shapeRange[i].PictureFormat.CropTop += CropLabSettings.GetAnchorY() * heightToCrop / scaleHeight;
                    shapeRange[i].PictureFormat.CropLeft += CropLabSettings.GetAnchorX() * widthToCrop / scaleWidth;
                    shapeRange[i].PictureFormat.CropRight += (1 - CropLabSettings.GetAnchorX()) * widthToCrop / scaleWidth;
                    shapeRange[i].PictureFormat.CropBottom += (1 - CropLabSettings.GetAnchorY()) * heightToCrop / scaleHeight;
                }
            }
            return hasChange;
        }
    }
}
