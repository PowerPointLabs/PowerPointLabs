using System;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineArrowFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            try
            {
                SyncFormat(formatShape, formatShape);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.Line.BeginArrowheadLength = formatShape.Line.BeginArrowheadLength;
            newShape.Line.BeginArrowheadStyle = formatShape.Line.BeginArrowheadStyle;
            newShape.Line.BeginArrowheadWidth = formatShape.Line.BeginArrowheadWidth;

            newShape.Line.EndArrowheadLength = formatShape.Line.EndArrowheadLength;
            newShape.Line.EndArrowheadStyle = formatShape.Line.EndArrowheadStyle;
            newShape.Line.EndArrowheadWidth = formatShape.Line.EndArrowheadWidth;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddLine(
                0, SyncFormatConstants.DisplayImageSize.Height,
                SyncFormatConstants.DisplayImageSize.Width, 0);
            SyncFormat(formatShape, shape);
            Bitmap image = new Bitmap(Graphics.ShapeToImage(shape));
            shape.Delete();
            return image;
        }
    }
}
