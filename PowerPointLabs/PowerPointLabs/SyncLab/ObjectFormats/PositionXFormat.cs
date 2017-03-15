using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class PositionXFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.Left = formatShape.Left;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Left).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
