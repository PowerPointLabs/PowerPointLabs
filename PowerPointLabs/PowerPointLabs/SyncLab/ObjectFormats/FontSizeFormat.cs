using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontSizeFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.TextEffect.FontSize = formatShape.TextEffect.FontSize;
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
