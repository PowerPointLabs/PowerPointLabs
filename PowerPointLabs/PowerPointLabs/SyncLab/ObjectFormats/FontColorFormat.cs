using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontColorFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.TextFrame.TextRange.Font.Color.RGB = formatShape.TextFrame.TextRange.Font.Color.RGB;
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
