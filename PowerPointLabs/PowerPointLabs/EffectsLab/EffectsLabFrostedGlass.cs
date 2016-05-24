using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public class EffectsLabFrostedGlass
    {
        private static Models.PowerPointSlide _slide;

        private const string MessageBoxTitle = "Error";
        private const string ErrorMessageNoSelection = TextCollection.EffectsLabErrorFrostedGlassNoSelection;
        private const string ErrorMessageNonShapeOrTextBox = TextCollection.EffectsLabErrorFrostedGlassNonShapeOrTextBox;
        private const string ErrorMessageEmptyTextBox = TextCollection.EffectsLabErrorFrostedGlassEmptyTextBox;

        private static readonly string BlurPicture = Path.GetTempPath() + @"\blur.png";

        public static void FrostedGlassEffect(Models.PowerPointSlide slide, float slideWidth, float slideHeight,
            PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes
                && selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            FrostedGlassEffect(slide, slideWidth, slideHeight, selection.ShapeRange);
        }

        public static void FrostedGlassEffect(Models.PowerPointSlide slide, float slideWidth, float slideHeight,
            PowerPoint.ShapeRange shapeRange)
        {
            if (shapeRange.Count == 0)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            PowerPoint.Shape blurSlideShape = null;

            try
            {
                _slide = slide;

                shapeRange.Cut();
            
                Utils.Graphics.ExportSlide(_slide, BlurPicture);
                blurSlideShape = GetBlurShape(BlurPicture);
                FitToSlide.AutoFit(blurSlideShape, slideWidth, slideHeight);

                shapeRange = slide.Shapes.Paste();

                var ungroupedShapeRange = UngroupAllForFrostedGlassShapeRange(shapeRange);
                var textBoxes = new List<PowerPoint.Shape>();
                var blurShapeRange = GetFrostedGlassShapeRange(ungroupedShapeRange, ref textBoxes);

                blurSlideShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                var blurShape = CropToShape.Crop(blurShapeRange, isInPlace: true, handleError: false);
                blurSlideShape.Delete();
                blurSlideShape = null;

                var overlayShape = GetOverlayShape(blurShape);
                if (overlayShape.Type == Office.MsoShapeType.msoGroup)
                {
                    var overlayShapeRange = overlayShape.Ungroup();
                    overlayShapeRange.MergeShapes(Office.MsoMergeCmd.msoMergeUnion);
                }

                foreach (var textBox in textBoxes)
                {
                    textBox.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                    // prevent offset when cut and paste a shape that have another shape in the same location
                    textBox.IncrementLeft(-12);
                    textBox.IncrementTop(-12);
                }
            }
            catch (Exception e)
            {
                if (blurSlideShape != null)
                {
                    blurSlideShape.Delete();
                }

                ShowErrorMessageBox(e.Message, e);
            }
        }

        private static PowerPoint.ShapeRange UngroupAllForFrostedGlassShapeRange(PowerPoint.ShapeRange shapeRange)
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
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    if (String.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
                    {
                        throw new Exception(ErrorMessageEmptyTextBox);
                    }

                    ungroupedShapeNames.Add(shape.Name);
                }
                else if (shape.Type == Office.MsoShapeType.msoAutoShape
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

        private static PowerPoint.ShapeRange GetFrostedGlassShapeRange(PowerPoint.ShapeRange shapeRange,
            ref List<PowerPoint.Shape> textBoxes)
        {
            var blurShapeNames = new List<string>();
            var queue = new Queue<PowerPoint.Shape>();

            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    var textBoundaryShape = GetShapeFromTextBoundary(shape);

                    // prevent offset when cut and paste a shape that have another shape in the same location
                    shape.IncrementLeft(12);
                    shape.IncrementTop(12);

                    textBoxes.Add(shape);
                    blurShapeNames.Add(textBoundaryShape.Name);
                }
                else // if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    if (!String.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
                    {
                        var textBox = DuplicateShapeInPlace(shape);
                        textBox.Fill.Visible = Office.MsoTriState.msoFalse;
                        textBox.Line.Visible = Office.MsoTriState.msoFalse;

                        // prevent offset when cut and paste a shape that have another shape in the same location
                        textBox.IncrementLeft(12);
                        textBox.IncrementTop(12);

                        textBoxes.Add(textBox);
                        shape.TextFrame2.DeleteText();
                    }

                    blurShapeNames.Add(shape.Name);
                }
            }

            var frostedGlassShapeRange = _slide.Shapes.Range(blurShapeNames.ToArray());

            return frostedGlassShapeRange;
        }

        private static PowerPoint.Shape GetBlurShape(string imageFile)
        {
            using (var imageFactory = new ImageProcessor.ImageFactory())
            {
                var image = imageFactory
                    .Load(imageFile)
                    .Image;

                var ratio = (float)image.Width / image.Height;
                var targetHeight = Math.Round(1100f - (1100f - 11f) / 100f * 95);
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

            var blurShape = _slide.Shapes.AddPicture(imageFile, Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue, 0, 0);

            return blurShape;
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
                || content.Equals(ErrorMessageNonShapeOrTextBox)
                || content.Equals(ErrorMessageEmptyTextBox))
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
