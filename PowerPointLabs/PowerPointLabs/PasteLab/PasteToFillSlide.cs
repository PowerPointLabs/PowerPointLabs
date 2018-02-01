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

            // Add code to compress the slide here, using ShapeToBitmap method from GraphicsUtil.cs
            System.Drawing.Bitmap shapeBitMap = GraphicsUtil.ShapeToBitmap(shapeToFillSlide);
            if (((System.Drawing.Image)shapeBitMap).HorizontalResolution > 96.0f && ((System.Drawing.Image)shapeBitMap).VerticalResolution > 96.0f)
            {
                shapeBitMap.SetResolution(96.0f, 96.0f);
            }
            // Comvert bitmap back into shape


            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);

            ppShapeToFillSlide.AbsoluteHeight = slideHeight;

            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);
        }
        public static PPShape Resize(PPShape originalShape, float w, float h)
        {
            //Original Image attributes
            float originalWidth = originalShape.AbsoluteWidth;
            float originalHeight = originalShape.AbsoluteHeight;

            // Figure out the ratio
            double ratioX = (double)w / (double)originalWidth;
            double ratioY = (double)h / (double)originalHeight;
            // use whichever multiplier is smaller
            double ratio = ratioX < ratioY ? ratioX : ratioY;

            // now we can get the new height and width
            int newHeight = System.Convert.ToInt32(originalHeight * ratio);
            int newWidth = System.Convert.ToInt32(originalWidth * ratio);

            originalShape.AbsoluteWidth = newWidth;
            originalShape.AbsoluteHeight = newHeight;

            return originalShape;
            /*
            Image thumbnail = new System.Drawing.Bitmap(newWidth, newHeight);
            Graphics graphic = System.Drawing.Graphics.FromImage(thumbnail);

            graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphic.SmoothingMode = SmoothingMode.HighQuality;
            graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphic.CompositingQuality = CompositingQuality.HighQuality;

            graphic.Clear(Color.Transparent);
            graphic.DrawImage(originalImage, 0, 0, newWidth, newHeight);

            return thumbnail;
            */
        }

    }
}
