using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FillFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Fill");
            }
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

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPatterned)
                {
                    newShape.Fill.Patterned(formatShape.Fill.Pattern);
                    newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                    newShape.Fill.BackColor.RGB = formatShape.Fill.BackColor.RGB;
                }
                else if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillBackground)
                {
                    newShape.Fill.Background();
                }
                else if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                {
                    newShape.Fill.Solid();
                    newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
