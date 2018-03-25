using System.Drawing;
using System.IO;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
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

            Shape pastingShape = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                pastingShape = pastingShapes.Group();
            }

            // Compression of large image(s)
            Shape shapeToFillSlide = GraphicsUtil.CompressImageInShape(pastingShape, slide);

            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            
            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);
            ppShapeToFillSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);

            try
            {
                shapeToFillSlide.Select();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // do nothing, let user select
            }
        }
    }
}
