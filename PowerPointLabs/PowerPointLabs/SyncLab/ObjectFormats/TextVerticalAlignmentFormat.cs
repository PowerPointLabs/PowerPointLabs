using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class TextVerticalAlignmentFormat : Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Vertical Text Alignment");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            string alignmentArt =
                "||\n" +
                "||\n";
            return SyncFormatUtil.GetTextDisplay(
                alignmentArt,
                new System.Drawing.Font(formatShape.TextEffect.FontName,
                                        SyncFormatConstants.DisplayImageFontSize,
                                        FontStyle.Bold),
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextFrame.VerticalAnchor = formatShape.TextFrame2.VerticalAnchor;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync TextAlignmentFormat");
                return false;
            }
            return true;
        }
    }
}
