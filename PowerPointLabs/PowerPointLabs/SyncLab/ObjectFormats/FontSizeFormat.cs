using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontSizeFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return SyncFormat(formatShape, formatShape);
        }

        public static bool SyncFormat(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextEffect.FontSize = formatShape.TextEffect.FontSize;
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
                Math.Round(formatShape.TextEffect.FontSize).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
