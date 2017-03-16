using System;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFillFormat
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
            newShape.Line.ForeColor = formatShape.Line.ForeColor;
            newShape.Line.BackColor = formatShape.Line.BackColor;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0,
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            SyncFormat(formatShape, shape);
            Bitmap image = new Bitmap(Graphics.ShapeToImage(shape));
            shape.Delete();
            return image;
        }
    }
}
