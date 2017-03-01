using System.Linq;
using System.Text.RegularExpressions;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    internal class CropToAspectRatio
    {
        public static PowerPoint.ShapeRange Crop(PowerPoint.Selection selection, float aspectRatio)
        {
            var croppedShape = Crop(selection.ShapeRange, aspectRatio);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }
            return croppedShape;
        }

        public static PowerPoint.ShapeRange Crop(PowerPoint.ShapeRange shapeRange, float aspectRatio)
        {
            for (int i = 1; i <= shapeRange.Count; i++)
            {
                PowerPoint.ShapeRange origShape = shapeRange[i].Duplicate();
                origShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                origShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                float origWidth = origShape.Width;
                float origHeight = origShape.Height;
                origShape.Delete();

                float currentWidth = shapeRange[i].Width - (shapeRange[i].PictureFormat.CropLeft + shapeRange[i].PictureFormat.CropRight) / origWidth;
                float currentHeight = shapeRange[i].Height - (shapeRange[i].PictureFormat.CropTop + shapeRange[i].PictureFormat.CropBottom) / origHeight;
                float currentProportions = currentWidth / currentHeight;

                if (currentProportions > aspectRatio)
                {
                    // Crop the width
                    float desiredWidth = currentHeight * aspectRatio;
                    float widthToCropEachSide = (currentWidth - desiredWidth) / 2.0f;
                    float widthToCropEachSideRatio = widthToCropEachSide / currentWidth;
                    shapeRange[i].PictureFormat.CropLeft += origWidth * widthToCropEachSideRatio;
                    shapeRange[i].PictureFormat.CropRight += origWidth * widthToCropEachSideRatio;
                }
                else if (currentProportions < aspectRatio)
                {
                    // Crop the height
                    float desiredHeight = currentWidth / aspectRatio;
                    float heightToCropEachSide = (currentHeight - desiredHeight) / 2.0f;
                    float heightToCropEachSideRatio = heightToCropEachSide / currentHeight;
                    shapeRange[i].PictureFormat.CropTop += origHeight * heightToCropEachSideRatio;
                    shapeRange[i].PictureFormat.CropBottom += origHeight * heightToCropEachSideRatio;
                }
            }

            return shapeRange;
        }
    }
}
