using System;
using System.Drawing;
using System.Drawing.Text;

using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab
{
    public class SyncFormatUtil
    {
        #region Display Image Utils

        public static Shapes GetTemplateShapes()
        {
            SyncLabShapeStorage shapeStorage = SyncLabShapeStorage.Instance;
            return shapeStorage.GetTemplateShapes();
        }
        
        public static Bitmap GetTextDisplay(string text, System.Drawing.Font font, Size imageSize)
        {
            Bitmap image = new Bitmap(imageSize.Width, imageSize.Height);
            Graphics g = Graphics.FromImage(image);
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            SizeF textSize = g.MeasureString(text, font);
            if (textSize.Width == 0 || textSize.Height == 0)
            {
                // nothing to print
                return image;
            }
            if (textSize.Width > imageSize.Width || textSize.Height > imageSize.Height)
            {
                double scale = Math.Min(imageSize.Width / textSize.Width, imageSize.Height / textSize.Height);
                font = new System.Drawing.Font(font.FontFamily, Convert.ToSingle(font.Size * scale),
                                                            font.Style, font.Unit, font.GdiCharSet, font.GdiVerticalFont);
                textSize = g.MeasureString(text, font);
            }
            float xPos = Convert.ToSingle((imageSize.Width - textSize.Width) / 2);
            float yPos = Convert.ToSingle((imageSize.Height - textSize.Height) / 2);
            g.DrawString(text, font, Brushes.Black, xPos, yPos);
            g.Dispose();
            return image;
        }

        #endregion

        #region Shape Name Utils
        public static bool IsValidFormatName(string name)
        {
            name = name.Trim();
            return name.Length > 0;
        }

        #endregion

    }
}
