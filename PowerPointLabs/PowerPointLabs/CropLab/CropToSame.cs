using System;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    public class CropToSame
    {

        private const float Epsilon = 0.00001F; // Prevents divide by zero

        public static void CropSelection(PowerPoint.ShapeRange shapeRange)
        {
            var refShape = shapeRange[1];
            float refScaleWidth = Graphics.GetScaleWidth(refShape);
            float refScaleHeight = Graphics.GetScaleHeight(refShape);

            for (int i = 2; i <= shapeRange.Count; i++)
            {
                float scaleWidth = Graphics.GetScaleWidth(shapeRange[i]);
                float scaleHeight = Graphics.GetScaleHeight(shapeRange[i]);
                float heightToCrop = shapeRange[i].Height - refShape.Height;
                float widthToCrop = shapeRange[i].Width - refShape.Width;

                float cropTop = Math.Max(shapeRange[1].PictureFormat.CropTop, Epsilon);
                float cropBottom = Math.Max(shapeRange[1].PictureFormat.CropBottom, Epsilon);
                float cropLeft = Math.Max(shapeRange[1].PictureFormat.CropLeft, Epsilon);
                float cropRight = Math.Max(shapeRange[1].PictureFormat.CropRight, Epsilon);

                float refShapeCroppedHeight = cropTop + cropBottom;
                float refShapeCroppedWidth = cropLeft + cropRight;

                shapeRange[i].PictureFormat.CropTop = Math.Max(0, heightToCrop * cropTop / refShapeCroppedHeight / scaleHeight);
                shapeRange[i].PictureFormat.CropLeft = Math.Max(0, widthToCrop * cropLeft / refShapeCroppedWidth / scaleWidth);
                shapeRange[i].PictureFormat.CropRight = Math.Max(0, widthToCrop * cropRight / refShapeCroppedWidth / scaleWidth);
                shapeRange[i].PictureFormat.CropBottom = Math.Max(0, heightToCrop * cropBottom / refShapeCroppedHeight / scaleHeight);
            }
        }
    }
}
