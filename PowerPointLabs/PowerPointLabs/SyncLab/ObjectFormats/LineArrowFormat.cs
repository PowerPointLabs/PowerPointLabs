using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineArrowFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            newShape.Line.BeginArrowheadLength = formatShape.Line.BeginArrowheadLength;
            newShape.Line.BeginArrowheadStyle = formatShape.Line.BeginArrowheadStyle;
            newShape.Line.BeginArrowheadWidth = formatShape.Line.BeginArrowheadWidth;

            newShape.Line.EndArrowheadLength = formatShape.Line.EndArrowheadLength;
            newShape.Line.EndArrowheadStyle = formatShape.Line.EndArrowheadStyle;
            newShape.Line.EndArrowheadWidth = formatShape.Line.EndArrowheadWidth;
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
