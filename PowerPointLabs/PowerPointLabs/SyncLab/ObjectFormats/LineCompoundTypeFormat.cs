using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineCompoundTypeFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Line Style");
            }
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddLine(
                0, SyncFormatConstants.DisplayImageSize.Height,
                SyncFormatConstants.DisplayImageSize.Width, 0);
            SyncFormat(formatShape, shape);
            shape.Line.ForeColor.RGB = SyncFormatConstants.ColorBlack;
            shape.Line.Weight = SyncFormatConstants.DisplayLineWeight;
            Bitmap image = Graphics.ShapeToBitmap(shape);
            shape.Delete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.Style = formatShape.Line.Style;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
