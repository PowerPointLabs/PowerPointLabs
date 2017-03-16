using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontFormat
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
                newShape.TextEffect.FontName = formatShape.TextEffect.FontName;
            }
            catch (Exception)
            {
                Logger.Log(newShape.Type + " unable to sync Font Format");
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
