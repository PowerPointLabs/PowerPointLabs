
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteToFillSlide
    {
        public static void Execute(PowerPointPresentation pres, PowerPointSlide slide, ShapeRange pastingShapes, float slideWidth, float slideHeight)
        {
            pastingShapes = ShapeUtil.GetShapesWhenTypeNotMatches(slide, pastingShapes, Microsoft.Office.Core.MsoShapeType.msoPlaceholder);
            if (pastingShapes.Count == 0)
            {
                return;
            }

            Shape pastingShape = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                pastingShape = pastingShapes.Group();
            }

            // Temporary house the latest clipboard shapes
            ShapeRange origClipboardShapes = ClipboardUtil.PasteShapesFromClipboard(pres, slide);
            // Compression of large image(s)
            Shape shapeToFillSlide = GraphicsUtil.CompressImageInShape(pastingShape, slide);
            // Bring the same original shapes back into clipboard, preserving original size
            origClipboardShapes.Cut();

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
