using System;
using System.Drawing;
using System.Drawing.Text;

using Microsoft.Office.Interop.PowerPoint;

using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.SyncLab
{
    public class SyncFormatUtil
    {
        #region Display Image Utils

        public static Shapes GetTemplateShapes()
        {
            Design design = Graphics.GetDesign(TextCollection.SyncLabStorageTemplateName);
            if (design == null)
            {
                design = Graphics.CreateDesign(TextCollection.SyncLabStorageTemplateName);
            }
            return design.TitleMaster.Shapes;
        }

        public static Bitmap GetTextDisplay(string text, System.Drawing.Font font, Size imageSize)
        {
            Bitmap image = new Bitmap(imageSize.Width, imageSize.Height);
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(image);
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            SizeF textSize = g.MeasureString(text, font);
            if (textSize.Width > imageSize.Width || textSize.Height > imageSize.Height)
            {
                double scale = Math.Min(imageSize.Width / textSize.Width, imageSize.Height / textSize.Height);
                System.Drawing.Font newFont = new System.Drawing.Font(font.FontFamily, Convert.ToSingle(font.Size * scale),
                                                            font.Style, font.Unit, font.GdiCharSet, font.GdiVerticalFont);
                return GetTextDisplay(text, newFont, imageSize);
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
