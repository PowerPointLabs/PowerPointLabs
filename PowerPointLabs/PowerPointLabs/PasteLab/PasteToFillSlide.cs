using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteToFillSlide
    {
        public static void Execute(PowerPointSlide slide, ShapeRange pastingShapes, float slideWidth, float slideHeight)
        {
            pastingShapes = ShapeUtil.GetShapesWhenTypeNotMatches(slide, pastingShapes, Microsoft.Office.Core.MsoShapeType.msoPlaceholder);
            if (pastingShapes.Count == 0)
            {
                return;
            }

            Shape shapeToFillSlide = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                shapeToFillSlide = pastingShapes.Group();
            }
            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);

            ppShapeToFillSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);

            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);
        }
    }
}
