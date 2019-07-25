using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShortcutsLab
{
    internal static class FillSlide
    {
        public static void Fill(PowerPoint.Selection selection, PowerPointSlide slide, float slideWidth, float slideHeight)
        {
            // Obtain selection, EnabledHandler has already checked if the selection are shapes/pics
            PowerPoint.Shape shapeToFitSlide = GetShapeFromSelection(slide, selection);
            
            // Fill operation
            shapeToFitSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            PPShape ppShapeToFitSlide = new PPShape(shapeToFitSlide);
            
            ppShapeToFitSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFitSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFitSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFitSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);

            CropLab.CropToSlide.Crop(shapeToFitSlide, slide, slideWidth, slideHeight);
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPointSlide slide, PowerPoint.Selection selection)
        {
            PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
            PowerPoint.Shape result = shapeRange.SafeGroup(slide);
            return result;
        }
    }
}
