using System;
using System.Drawing;
using System.Drawing.Text;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.SyncLab.Views;

namespace PowerPointLabs.SyncLab
{
    public class SyncFormatUtil
    {
        #region Display Image Utils

        public static Shapes GetTemplateShapes()
        {
            SyncLabShapeStorage shapeStorage = SyncLabShapeStorage.Instance;
            return shapeStorage.Slides[SyncLabShapeStorage.FormatStorageSlide].Shapes;
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

        #region Sync Shape Format utils

        /// <summary>
        /// Applies the specified formats from one shape to multiple shapes
        /// </summary>
        /// <param name="nodes">styles to apply</param>
        /// <param name="formatShape">source shape</param>
        /// <param name="newShapes">destination shape</param>
        public static void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, ShapeRange newShapes)
        {
            foreach (Shape newShape in newShapes)
            {
                ApplyFormats(nodes, formatShape, newShape);
            }
        }

        public static void ApplyFormats(FormatTreeNode[] nodes, Shape formatShape, Shape newShape)
        {
            foreach (FormatTreeNode node in nodes)
            {
                ApplyFormats(node, formatShape, newShape);
            }
        }

        public static void ApplyFormats(FormatTreeNode node, Shape formatShape, Shape newShape)
        {
            if (node.Format != null)
            {
                if (!node.IsChecked.HasValue || !node.IsChecked.Value)
                {
                    return;
                }
                node.Format.SyncFormat(formatShape, newShape);
            }
            else
            {
                ApplyFormats(node.ChildrenNodes, formatShape, newShape);
            }
        }
        #endregion
    }
}
