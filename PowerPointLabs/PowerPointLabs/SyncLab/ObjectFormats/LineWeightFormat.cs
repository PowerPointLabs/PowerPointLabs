using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineWeightFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return SyncFormat(formatShape, formatShape);
        }

        public static bool SyncFormat(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.Weight = formatShape.Line.Weight;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(Math.Max(formatShape.Line.Weight, 0)).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
