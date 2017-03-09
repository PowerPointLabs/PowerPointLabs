using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;


namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontColorFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.TextFrame.TextRange.Font.Color.RGB = formatShape.TextFrame.TextRange.Font.Color.RGB;
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Bitmap b = new Bitmap(200, 200);
            Graphics g = Graphics.FromImage(b);
            g.FillRectangle(Brushes.DarkBlue, 0, 0, 200, 200);
            return b;
        }
    }
}
