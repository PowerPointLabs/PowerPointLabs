using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineTransparencyFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Transparency");
            }
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Line.Transparency * 100).ToString() + "%",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.Transparency = formatShape.Line.Transparency;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
