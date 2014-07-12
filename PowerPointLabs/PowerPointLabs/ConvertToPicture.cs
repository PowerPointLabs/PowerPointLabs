using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class ConvertToPicture
    {
        public static void Convert(PowerPoint.Selection selection)
        {
            if (IsSelectionShape(selection))
            {
                var shape = GetShapeFromSelection(selection);
                shape = CutPasteShape(shape);
                ConvertToPictureForShape(shape);
            }
            else
            {
                MessageBox.Show(TextCollection.ErrorTypeNotSupported, TextCollection.ErrorWindowTitle);
            }
        }

        public static void ConvertAndSave(PowerPoint.Selection selection, string fileName)
        {
            if (IsSelectionShape(selection))
            {
                var grouped = selection.ShapeRange.Count > 1;

                var shape = GetShapeFromSelection(selection);
                shape = CutPasteShape(shape);
                shape.Export(fileName, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0,
                             PowerPoint.PpExportMode.ppScaleXY);

                if (grouped)
                {
                    shape.Ungroup().Select();
                }
                else
                {
                    shape.Select();
                }
            }
            else
            {
                MessageBox.Show(TextCollection.ErrorTypeNotSupported, TextCollection.ErrorWindowTitle);
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
                PowerPointLabsGlobals.LogException(e, "Chart cannot be rotated.");
            }
            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            shape.Delete();
            var pic = PowerPointLabsGlobals.GetCurrentSlide().Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x + (width - pic.Width) / 2;
            pic.Top = y + (height - pic.Height) / 2;
            pic.Rotation = rotation;
            pic.Select();
        }

        public static System.Drawing.Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ConvertToPicture);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetConvertToPicMenuImage");
                throw;
            }
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
            shape = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste()[1];
            return shape;
        }

        private static PowerPoint.Shape GetShapeFromSelection(PowerPoint.Selection selection)
        {
            PowerPoint.Shape shape = 
                selection.ShapeRange.Count > 1 ? selection.ShapeRange.Group() : selection.ShapeRange[1];
            return shape;
        }

        private static bool IsSelectionShape(PowerPoint.Selection selection)
        {
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes;
        }
    }
}
