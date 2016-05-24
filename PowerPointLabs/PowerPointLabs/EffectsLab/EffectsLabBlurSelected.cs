using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.EffectsLab
{
    public class EffectsLabBlurSelected
    {
        private static Models.PowerPointSlide _slide;

        private const string MessageBoxTitle = "Error";
        private const string ErrorMessageNoSelection = TextCollection.EffectsLabBlurSelectedErrorNoSelection;
        private const string ErrorMessageNonShapeOrTextBox = TextCollection.EffectsLabBlurSelectedErrorNonShapeOrTextBox;

        private static readonly string BlurPicture = Path.GetTempPath() + @"\blur.png";

        public static void BlurSelected(Models.PowerPointSlide slide, float slideWidth, float slideHeight,
            PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes
                && selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            BlurSelected(slide, slideWidth, slideHeight, selection.ShapeRange);
        }

        public static void BlurSelected(Models.PowerPointSlide slide, float slideWidth, float slideHeight,
            PowerPoint.ShapeRange shapeRange)
        {
            if (shapeRange.Count == 0)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                _slide = slide;

                shapeRange.Cut();
            
                Utils.Graphics.ExportSlide(_slide, BlurPicture);
                BlurImage(BlurPicture, 95);

                shapeRange = slide.Shapes.Paste();

                var ungroupedRange = UngroupAllShapeRange(shapeRange);
                var shapeGroups = GetShapeGroups(ungroupedRange);

                foreach (var shapes in shapeGroups)
                {
                    if (shapes.Length == 2)
                    {
                        shapes[0].ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                        // prevent offset when cut and paste a shape that have another shape in the same location
                        shapes[0].IncrementLeft(-12);
                        shapes[0].IncrementTop(-12);

                        var range = _slide.Shapes.Range(new[] { shapes[0].Name, shapes[1].Name });
                        range.Group();
                    }
                }
            }
            catch (Exception e)
            {
                ActionFramework.Common.Log.Logger.LogException(e, "BlurSelectedEffect");

                ShowErrorMessageBox(e.Message, e);
            }
        }

        private static PowerPoint.ShapeRange UngroupAllShapeRange(PowerPoint.ShapeRange shapeRange)
        {
            var ungroupedShapeNames = new List<string>();
            var queue = new Queue<PowerPoint.Shape>();

            foreach (PowerPoint.Shape shape in shapeRange)
            {
                queue.Enqueue(shape);
            }

            while (queue.Count != 0)
            {
                var shape = queue.Dequeue();

                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (PowerPoint.Shape item in subRange)
                    {
                        queue.Enqueue(item);
                    }
                }
                else if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox
                    || shape.Type == Office.MsoShapeType.msoAutoShape
                    || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    ungroupedShapeNames.Add(shape.Name);
                }
                else
                {
                    throw new Exception(ErrorMessageNonShapeOrTextBox);
                }
            }

            var ungroupedShapeRange = _slide.Shapes.Range(ungroupedShapeNames.ToArray());

            return ungroupedShapeRange;
        }

        private static PowerPoint.Shape[][] GetShapeGroups(PowerPoint.ShapeRange shapeRange)
        {
            var shapeGroups = new PowerPoint.Shape[shapeRange.Count][];

            for (int i = 0; i < shapeRange.Count; i++)
            {
                var shape = shapeRange[i + 1];

                if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    var textBoundaryShape = GetShapeFromTextBoundary(shape);
                    textBoundaryShape = CropToShape.FillInShapeWithImage(BlurPicture, textBoundaryShape, isInPlace: true);

                    // prevent offset when cut and paste a shape that have another shape in the same location
                    shape.IncrementLeft(12);
                    shape.IncrementTop(12);

                    shapeGroups[i] = new PowerPoint.Shape[] { shape, textBoundaryShape };
                }
                else // if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    if (!string.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
                    {
                        var textBox = DuplicateShapeInPlace(shape);
                        textBox.Fill.Visible = Office.MsoTriState.msoFalse;
                        textBox.Line.Visible = Office.MsoTriState.msoFalse;

                        // prevent offset when cut and paste a shape that have another shape in the same location
                        textBox.IncrementLeft(12);
                        textBox.IncrementTop(12);
                        
                        shape.TextFrame2.DeleteText();
                        shape = CropToShape.FillInShapeWithImage(BlurPicture, shape, isInPlace: true);

                        shapeGroups[i] = new PowerPoint.Shape[] { textBox, shape };
                    }
                    else
                    {
                        shape = CropToShape.FillInShapeWithImage(BlurPicture, shape, isInPlace: true);
                        shapeGroups[i] = new PowerPoint.Shape[] { shape };
                    }
                }
            }

            return shapeGroups;
        }

        private static void BlurImage(string imageFile, int degree)
        {
            if (degree != 0)
            {
                float originalWidth, originalHeight;

                using (var imageFactory = new ImageProcessor.ImageFactory())
                {
                    var image = imageFactory
                        .Load(imageFile)
                        .Image;

                    originalWidth = image.Width;
                    originalHeight = image.Height;

                    var ratio = (float)image.Width / image.Height;
                    var targetHeight = Math.Round(1100f - (1100f - 11f) / 100f * degree);
                    var targetWidth = Math.Round(targetHeight * ratio);

                    image = imageFactory
                        .Resize(new Size((int)targetWidth, (int)targetHeight))
                        .Image;
                    image.Save(imageFile);
                }

                using (var imageFactory = new ImageProcessor.ImageFactory())
                {
                    var image = imageFactory
                        .Load(imageFile)
                        .GaussianBlur(5)
                        .Image;
                    image.Save(imageFile);
                }

                using (var imageFactory = new ImageProcessor.ImageFactory())
                {
                    var image = imageFactory
                        .Load(imageFile)
                        .Resize(new Size((int)originalWidth, (int)originalHeight))
                        .Image;
                    image.Save(imageFile);
                }
            }
        }

        private static PowerPoint.Shape GetOverlayShape(PowerPoint.Shape shape)
        {
            var overlayShape = DuplicateShapeInPlace(shape);

            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Utils.StringUtil.GetColorFromHexValue("#000000"));
            overlayShape.Fill.Transparency = 80f / 100;
            overlayShape.Line.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Utils.StringUtil.GetColorFromHexValue("#000000"));
            overlayShape.Line.Transparency = 80f / 100;
            overlayShape.Line.Weight = 5;
            overlayShape.Line.Visible = Office.MsoTriState.msoFalse;

            return overlayShape;
        }

        private static PowerPoint.Shape GetShapeFromTextBoundary(PowerPoint.Shape shape)
        {
            var rotation = shape.Rotation;
            if (rotation != 0)
            {
                shape.Rotation = 0;
            }

            var textFrame = shape.TextFrame2;
            var textRange = textFrame.TextRange.TrimText();

            var left = textRange.BoundLeft - textFrame.MarginLeft;
            var top = textRange.BoundTop - textFrame.MarginTop;
            var width = textRange.BoundWidth + textFrame.MarginLeft + textFrame.MarginRight;
            var height = textRange.BoundHeight + textFrame.MarginTop + textFrame.MarginBottom;

            var textBoundaryShape = _slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);

            shape.Fill.Visible = Office.MsoTriState.msoFalse;
            shape.Line.Visible = Office.MsoTriState.msoFalse;

            // for placeholders with text out of box affected due to cut and paste
            // and text boxes with text out of box
            if (shape.Height < textBoundaryShape.Height)
            {
                shape.Left = textBoundaryShape.Left;
                shape.Top = textBoundaryShape.Top;
                shape.Width = textBoundaryShape.Width;
                shape.Height = textBoundaryShape.Height;
            }

            if (rotation != 0)
            {
                shape.Rotation = rotation;

                var origin = Utils.Graphics.GetCenterPoint(shape);
                var unrotatedCenter = Utils.Graphics.GetCenterPoint(textBoundaryShape);
                var rotatedCenter = Utils.Graphics.RotatePoint(unrotatedCenter, origin, rotation);

                textBoundaryShape.Left += (rotatedCenter.X - unrotatedCenter.X);
                textBoundaryShape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

                textBoundaryShape.Rotation = PositionsLab.PositionsLabMain.AddAngles(textBoundaryShape.Rotation, rotation);
            }

            return textBoundaryShape;
        }

        private static PowerPoint.Shape DuplicateShapeInPlace(PowerPoint.Shape shape)
        {
            var duplicateShape = shape.Duplicate()[1];
            duplicateShape.Left = shape.Left;
            duplicateShape.Top = shape.Top;

            var match = System.Text.RegularExpressions.Regex.Match(duplicateShape.Name, @"\d+$");
            if (!match.Success || int.Parse(match.Value) != duplicateShape.Id - 1)
            {
                duplicateShape.Name += " " + (duplicateShape.Id - 1);
            }

            return duplicateShape;
        }

        private static void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception == null
                || content.Equals(ErrorMessageNonShapeOrTextBox))
            {
                MessageBox.Show(content, MessageBoxTitle);
            }
            else
            {
                Views.ErrorDialogWrapper.ShowDialog(MessageBoxTitle, content, exception);
            }
        }
    }
}
