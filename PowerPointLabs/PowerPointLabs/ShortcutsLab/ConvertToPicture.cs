using System;
using System.Collections.Generic;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
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

        public static void Convert(PowerPointPresentation pres, PowerPointSlide slide, PowerPoint.Selection selection)
        {
            if (ShapeUtil.IsSelectionShapeOrText(selection))
            {
                PowerPoint.Shape shape = GetShapeFromSelection(selection);
                int originalZOrder = shape.ZOrderPosition;
                // In case shape is corrupted
                if (ShapeUtil.IsCorrupted(shape))
                {
                    shape = ShapeUtil.CorruptionCorrection(shape, slide);
                }
                ConvertToPictureForShape(pres, slide, shape, originalZOrder);
            }
            else
            {
                MessageBox.Show(ShortcutsLabText.ErrorTypeNotSupported, ShortcutsLabText.ErrorWindowTitle);
            }
        }

        public static bool ConvertAndSave(ShapeRange selectedShapes, string fileName)
        {
            if (!ShapeUtil.IsShapeRangeShapeOrText(selectedShapes))
            {
                MessageBox.Show(ShortcutsLabText.ErrorTypeNotSupported, ShortcutsLabText.ErrorWindowTitle);
                return false;
            }

            try
            {
                GraphicsUtil.ExportShape(selectedShapes, fileName);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception during export shapes: " + e.Message, ShortcutsLabText.ErrorWindowTitle);
                return false;
            }
        }

        private static void ConvertToPictureForShape(PowerPointPresentation pres, PowerPointSlide slide, PowerPoint.Shape shape, int originalZOrder)
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

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                shape.Copy();
                float x = shape.Left;
                float y = shape.Top;
                float width = shape.Width;
                float height = shape.Height;
                shape.Delete();
                PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
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
                return pic;
            }, pres, slide);
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.Selection selection)
        {
            ShapeRange shapeRange = GetShapeRangeFromSelection(selection);
            Shape shape = GetShapeFromShapeRange(shapeRange);
            return shape;
        }

        private static ShapeRange GetShapeRangeFromSelection(Selection selection)
        {
            return selection.ShapeRange;
        }

        private static PowerPoint.Shape GetShapeFromShapeRange(PowerPoint.ShapeRange shapeRange)
        {
            return shapeRange.Count > 1 ? shapeRange.Group() : shapeRange[1];
        }

    }
}
