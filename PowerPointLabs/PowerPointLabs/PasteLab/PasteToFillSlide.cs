using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteToFillSlide
    {
        public static void Execute(PowerPointSlide slide, ShapeRange pastingShapes, float width, float height)
        {
            pastingShapes = Graphics.GetShapesWhenTypeNotMatches(slide, pastingShapes, Microsoft.Office.Core.MsoShapeType.msoPlaceholder);

            Shape shapeToFillSlide = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                shapeToFillSlide = pastingShapes.Group();
            }
            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);

            ppShapeToFillSlide.AbsoluteHeight = height;
            if (ppShapeToFillSlide.AbsoluteWidth < width)
            {
                ppShapeToFillSlide.AbsoluteWidth = width;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(width / 2, height / 2);

            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, width, height);
        }
    }
}
