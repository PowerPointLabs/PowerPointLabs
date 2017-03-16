using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontSizeFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            Sync(formatShape, newShape);
        }

        public static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextEffect.FontSize = formatShape.TextEffect.FontSize;
            }
            catch (Exception)
            {
                Logger.Log(newShape.Type + " unable to sync Font Size");
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
