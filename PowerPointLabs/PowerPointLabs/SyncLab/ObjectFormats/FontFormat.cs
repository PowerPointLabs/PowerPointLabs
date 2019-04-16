using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Font Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                SyncFormatConstants.DisplayFontString,
                new System.Drawing.Font(formatShape.TextEffect.FontName,
                                        SyncFormatConstants.DisplayImageFontSize),
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextEffect.FontName = formatShape.TextEffect.FontName;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync FontFormat");
                return false;
            }
        }
    }
}
