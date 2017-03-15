using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab
{
    public class SyncFormatUtil
    {
        #region Display Image Utils

        public static Shapes GetTemplateShapes()
        {
            Design design = Graphics.GetDesign(TextCollection.StorageTemplateName);
            if (design == null)
            {
                design = Graphics.CreateDesign(TextCollection.StorageTemplateName);
            }
            return design.TitleMaster.Shapes;
        }

        public static Bitmap GetTextDisplay(string text, System.Drawing.Font font, Size size)
        {
            Bitmap image = new Bitmap(size.Width, size.Height);
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(image);
            SizeF textSize = g.MeasureString(text, font);
            Bitmap textImage = new Bitmap((int)Math.Ceiling(textSize.Width), (int)Math.Ceiling(textSize.Height));
            System.Drawing.Graphics g2 = System.Drawing.Graphics.FromImage(textImage);
            g2.DrawString(text, font, Brushes.Black, 0, 0);

            double scale = Math.Min(size.Width / textSize.Width, size.Height / textSize.Height);
            double newWidth = textSize.Width * scale;
            double newHeight = textSize.Height * scale;
            double newX = (size.Width - newWidth) / 2;
            double newY = (size.Height - newHeight) / 2;
            g.DrawImage(textImage, Convert.ToSingle(newX), Convert.ToSingle(newY),
                Convert.ToSingle(newWidth), Convert.ToSingle(newHeight));
            return image;
        }

        #endregion
    }
}
