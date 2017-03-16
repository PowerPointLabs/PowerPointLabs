using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return SyncFormat(formatShape, formatShape);
        }

        public static bool SyncFormat(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextEffect.FontName = formatShape.TextEffect.FontName;
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
                "T",
                new System.Drawing.Font(formatShape.TextEffect.FontName,
                                        SyncFormatConstants.DisplayImageFontSize),
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
