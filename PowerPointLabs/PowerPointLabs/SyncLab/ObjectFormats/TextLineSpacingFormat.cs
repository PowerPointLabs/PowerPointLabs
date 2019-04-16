using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class TextLineSpacingFormat : Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Line Spacing");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            string spacingArt =
                "--------\n" +
                "____\n";
            return SyncFormatUtil.GetTextDisplay(
                spacingArt,
                new System.Drawing.Font(formatShape.TextEffect.FontName,
                                        SyncFormatConstants.DisplayImageFontSize,
                                        FontStyle.Bold),
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter =
                    formatShape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter;
                newShape.TextFrame.TextRange.ParagraphFormat.SpaceBefore =
                    formatShape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore;
                newShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin =
                    formatShape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync TextLineSpacingFormat");
                return false;
            }
            return true;
        }
    }
}
