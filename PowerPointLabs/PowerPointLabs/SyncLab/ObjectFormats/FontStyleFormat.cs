using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontStyleFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Font Style");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            System.Drawing.Font font = SyncFormatConstants.DisplayImageFont;
            Microsoft.Office.Interop.PowerPoint.Font formatFont = formatShape.TextFrame.TextRange.Font;
            FontStyle style = 0;
            if (formatFont.Underline == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                style |= FontStyle.Underline;
            }
            if (formatFont.Bold == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                style |= FontStyle.Bold;
            }
            if (formatFont.Italic == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                style |= FontStyle.Italic;
            }
            font = new System.Drawing.Font(font.FontFamily, font.Size, style);
            return SyncFormatUtil.GetTextDisplay(SyncFormatConstants.DisplayFontString, font,
                                                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                SyncTextRange(formatShape.TextFrame.TextRange, newShape.TextFrame.TextRange);
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync FontStyleFormat");
                return false;
            }
        }

        private static void SyncTextRange(TextRange formatTextRange, TextRange newTextRange)
        {
            newTextRange.Font.Underline = formatTextRange.Font.Underline;
            newTextRange.Font.Bold = formatTextRange.Font.Bold;
            newTextRange.Font.Italic = formatTextRange.Font.Italic;
        }
    }
}
