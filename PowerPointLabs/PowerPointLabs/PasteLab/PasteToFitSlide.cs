﻿using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteToFitSlide
    {
        public static void Execute(PowerPointSlide slide, ShapeRange pastingShapes, float slideWidth, float slideHeight)
        {
            pastingShapes = ShapeUtil.GetShapesWhenTypeNotMatches(slide, pastingShapes, Microsoft.Office.Core.MsoShapeType.msoPlaceholder);
            if (pastingShapes.Count == 0)
            {
                return;
            }

            Shape shapeToFitSlide = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                shapeToFitSlide = pastingShapes.Group();
            }
            shapeToFitSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            PPShape ppShapeToFitSlide = new PPShape(shapeToFitSlide);

            ResizeShape(ppShapeToFitSlide, slideWidth, slideHeight);
            ppShapeToFitSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
        }

        public static void ResizeShape(PPShape ppShapeToFitSlide, float w, float h)
        {
            //Original PPShape attributes
            float originalWidth = ppShapeToFitSlide.AbsoluteWidth;
            float originalHeight = ppShapeToFitSlide.AbsoluteHeight;

            // Figure out the ratio
            double ratioX = (double)w / (double)originalWidth;
            double ratioY = (double)h / (double)originalHeight;
            // use whichever multiplier is smaller
            double ratio = ratioX < ratioY ? ratioX : ratioY;

            // Now we can get the new height and width
            float newHeight = originalHeight * (float)ratio;
            float newWidth = originalWidth * (float)ratio;

            // Resize the image accordingly to the slide
            ppShapeToFitSlide.AbsoluteHeight = newHeight;
            ppShapeToFitSlide.AbsoluteWidth = newWidth;
        }
    }
}
