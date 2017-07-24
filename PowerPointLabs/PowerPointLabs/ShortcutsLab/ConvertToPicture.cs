using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShortcutsLab
{
    internal static class ConvertToPicture
    {
#pragma warning disable 0618
        public static void Convert(PowerPoint.Selection selection)
        {
            if (ShapeUtil.IsSelectionShapeOrText(selection))
            {
                var shape = GetShapeFromSelection(selection);
                shape = CutPasteShape(shape);
                ConvertToPictureForShape(shape);
            }
            else
            {
                MessageBox.Show(ShortcutsLabText.ErrorTypeNotSupported, ShortcutsLabText.ErrorWindowTitle);
            }
        }

        public static void ConvertAndSave(PowerPoint.Selection selection, string fileName)
        {
            if (ShapeUtil.IsSelectionShapeOrText(selection))
            {
                if (selection.HasChildShapeRange)
                {
                    GraphicsUtil.ExportShape(selection.ChildShapeRange, fileName);
                }
                else
                {
                    GraphicsUtil.ExportShape(selection.ShapeRange, fileName);
                }
            }
            else
            {
                MessageBox.Show(ShortcutsLabText.ErrorTypeNotSupported, ShortcutsLabText.ErrorWindowTitle);
            }
        }

        private static void ConvertToPictureForShape(PowerPoint.Shape shape)
        {
            float rotation = 0;
            try
            {
                rotation = shape.Rotation;
                shape.Rotation = 0;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Chart cannot be rotated.");
            }
            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            shape.Delete();
            var pic = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x + (width - pic.Width) / 2;
            pic.Top = y + (height - pic.Height) / 2;
            pic.Rotation = rotation;
            pic.Select();
        }

        /// <summary>
        /// To avoid corrupted shape.
        /// Corrupted shape is produced when delete or cut a shape programmatically, but then users undo it.
        /// After that, most of operations on corrupted shapes will throw an exception.
        /// One solution for this is to re-allocate its memory: simply cut/copy and paste.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static PowerPoint.Shape CutPasteShape(PowerPoint.Shape shape)
        {
            shape.Cut();
            shape = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste()[1];
            return shape;
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.Selection selection)
        {
            return GetShapeFromSelection(selection.ShapeRange);
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.ShapeRange shapeRange)
        {
            PowerPoint.Shape result = shapeRange.Count > 1 ? shapeRange.Group() : shapeRange[1];
            return result;
        }
    }
}
