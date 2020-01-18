using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFillFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Line Fill");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0,
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Fill.ForeColor = formatShape.Line.ForeColor;
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.SafeDelete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.ForeColor.RGB = formatShape.Line.ForeColor.RGB;
                newShape.Line.BackColor.RGB = formatShape.Line.BackColor.RGB;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LineFillFormat");
                return false;
            }
        }
    }
}
