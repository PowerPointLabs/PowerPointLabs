using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFillFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            //TODO not working
            newShape.Line.ForeColor = newShape.Line.ForeColor;
            newShape.Line.BackColor = formatShape.Line.BackColor;
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
