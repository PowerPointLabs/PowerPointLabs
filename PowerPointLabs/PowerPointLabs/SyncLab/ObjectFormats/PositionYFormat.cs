using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class PositionYFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.Top = formatShape.Top;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Top).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
