using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontColorFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return SyncFormat(formatShape, formatShape);
        }

        public static bool SyncFormat(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextFrame.TextRange.Font.Color.RGB = formatShape.TextFrame.TextRange.Font.Color.RGB;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0,
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Fill.ForeColor.RGB = formatShape.TextFrame.TextRange.Font.Color.RGB;
            shape.Fill.BackColor.RGB = formatShape.TextFrame.TextRange.Font.Color.RGB;
            shape.Fill.Solid();
            Bitmap image = new Bitmap(Graphics.ShapeToImage(shape));
            shape.Delete();
            return image;
        }
    }
}
