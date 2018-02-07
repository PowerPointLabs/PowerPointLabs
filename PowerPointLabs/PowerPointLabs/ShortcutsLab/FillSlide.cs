using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShortcutsLab
{
    internal static class FillSlide
    {
        public static void Fill(PowerPoint.Selection selection, PowerPointSlide slide, float slideWidth, float slideHeight)
        {
            // Obtain selection, EnabledHandler has already checked if the selection are shapes/pics
            PowerPoint.Shape shapeToFitSlide = GetShapeFromSelection(selection);
            
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

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.Selection selection)
        {
            return GetShapeFromSelection(selection.ShapeRange);
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.ShapeRange shapeRange)
        {
            PowerPoint.Shape result = shapeRange.Count > 1 ? shapeRange.Group() : shapeRange[1];
            return result;
        }
    }
}
