using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontStyleFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            //What is the difference between TextFrame and TextFrame2?
            SyncTextRange(formatShape.TextFrame.TextRange, newShape.TextFrame.TextRange);
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Bitmap b = new Bitmap(200, 200);
            Graphics g = Graphics.FromImage(b);
            g.FillRectangle(Brushes.DarkBlue, 0, 0, 200, 200);
            return b;
        }

        private static void SyncTextRange(TextRange formatTextRange, TextRange newTextRange)
        {
            newTextRange.Font.Underline = formatTextRange.Font.Underline;
            newTextRange.Font.Bold = formatTextRange.Font.Bold;
            newTextRange.Font.Italic = formatTextRange.Font.Italic;
        }
    }
}
