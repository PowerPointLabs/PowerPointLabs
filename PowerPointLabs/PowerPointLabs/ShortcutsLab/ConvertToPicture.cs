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
                PowerPoint.Shape shape = GetShapeFromSelection(selection);
                int originalZOrder = shape.ZOrderPosition;
                // In case shape is corrupted
                shape = ShapeUtil.CutPasteShape(shape);
                ConvertToPictureForShape(shape, originalZOrder);
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

        private static void ConvertToPictureForShape(PowerPoint.Shape shape, int originalZOrder)
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

            // Save clipboard onto a temp slide, this is similar to code in PasteLabActionHandler.cs
            // Have no choice to use deprecated methods because ConvertToPicture does not use ActionFramework
            PowerPointPresentation presentation = PowerPointPresentation.Current;

            PowerPointSlide tempClipboardSlide = presentation.AddSlide();
            PowerPoint.ShapeRange tempClipboardShapes = ClipboardUtil.PasteShapesFromClipboard(tempClipboardSlide);

            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            shape.Delete();
            PowerPoint.Shape pic = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x + (width - pic.Width) / 2;
            pic.Top = y + (height - pic.Height) / 2;
            pic.Rotation = rotation;
            // move picture to original z-order
            while (pic.ZOrderPosition > originalZOrder)
            {
                pic.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }
            while (pic.ZOrderPosition < originalZOrder)
            {
                pic.ZOrder(Office.MsoZOrderCmd.msoBringForward);
            }
            pic.Select();

            // Revert clipboard
            if (tempClipboardShapes != null)
            {
                tempClipboardShapes.Copy();
            }
            tempClipboardSlide.Delete();
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
