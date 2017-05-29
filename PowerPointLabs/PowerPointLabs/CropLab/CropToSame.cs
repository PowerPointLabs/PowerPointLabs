using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    public class CropToSame
    {

        private const float Epsilon = 0.00001F; // Prevents divide by zero

        public static bool CropSelection(PowerPoint.ShapeRange shapeRange)
        {
            bool hasChange = false;

            Shape refObj = shapeRange[1];
            bool isRefObjShape = Graphics.IsShape(refObj);

            float refScaleWidth = Graphics.GetScaleWidth(refObj);
            float refScaleHeight = Graphics.GetScaleHeight(refObj);

            float cropTop = isRefObjShape ? Epsilon : Math.Max(refObj.PictureFormat.CropTop, Epsilon);
            float cropBottom = isRefObjShape ? Epsilon : Math.Max(refObj.PictureFormat.CropBottom, Epsilon);
            float cropLeft = isRefObjShape ? Epsilon : Math.Max(refObj.PictureFormat.CropLeft, Epsilon);
            float cropRight = isRefObjShape ? Epsilon : Math.Max(refObj.PictureFormat.CropRight, Epsilon);

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

                float scaleWidth = Graphics.GetScaleWidth(shapeRange[i]);
                float scaleHeight = Graphics.GetScaleHeight(shapeRange[i]);
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
