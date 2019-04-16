using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineTransparencyFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Transparency");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
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
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LineTransparencyFormat");
                return false;
            }
        }
    }
}
