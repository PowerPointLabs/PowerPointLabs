using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineArrowFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Line Arrow");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddLine(
                0, SyncFormatConstants.DisplayImageSize.Height,
                SyncFormatConstants.DisplayImageSize.Width, 0);
            SyncFormat(formatShape, shape);
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.Delete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.BeginArrowheadLength = formatShape.Line.BeginArrowheadLength;
                newShape.Line.BeginArrowheadStyle = formatShape.Line.BeginArrowheadStyle;
                newShape.Line.BeginArrowheadWidth = formatShape.Line.BeginArrowheadWidth;

                newShape.Line.EndArrowheadLength = formatShape.Line.EndArrowheadLength;
                newShape.Line.EndArrowheadStyle = formatShape.Line.EndArrowheadStyle;
                newShape.Line.EndArrowheadWidth = formatShape.Line.EndArrowheadWidth;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LineArrowFormat");
                return false;
            }
            return true;
        }
    }
}
