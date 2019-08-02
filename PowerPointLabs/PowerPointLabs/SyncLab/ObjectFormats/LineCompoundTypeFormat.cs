using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineCompoundTypeFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Line Style");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddLine(
                0, SyncFormatConstants.DisplayImageSize.Height,
                SyncFormatConstants.DisplayImageSize.Width, 0);
            SyncFormat(formatShape, shape);
            shape.Line.ForeColor.RGB = SyncFormatConstants.ColorBlack;
            shape.Line.Weight = SyncFormatConstants.DisplayLineWeight;
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.SafeDelete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.Style = formatShape.Line.Style;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LineCompoundTypeFormat");
                return false;
            }
        }
    }
}
