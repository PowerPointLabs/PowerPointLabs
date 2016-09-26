using ImageProcessor;
using ImageProcessor.Imaging;
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
        public static bool IsTintSelected = false;
        public static bool IsTintRemainder = false;
        public static bool IsTintBackground = false;

        private static Models.PowerPointSlide _slide;

        private const string HexColor = "#000000";
        private const float Transparency = 0.8f;

        private const string MessageBoxTitle = "Error";
        private const string ErrorMessageNoSelection = TextCollection.EffectsLabBlurSelectedErrorNoSelection;
        private const string ErrorMessageNonShapeOrTextBox = TextCollection.EffectsLabBlurSelectedErrorNonShapeOrTextBox;

        private static readonly string BlurPicture = Path.GetTempPath() + @"\blur.png";

        public static PowerPoint.ShapeRange BlurSelected(Models.PowerPointSlide slide, PowerPoint.Selection selection, int percentage)
        {
            if (!IsValidSelection(selection))
            {
                return null;
            }

            var range = BlurSelected(slide, selection.ShapeRange, percentage);
            if (range != null)
            {
                range.Select();
            }

            return range;
        }

        public static PowerPoint.ShapeRange BlurSelected(Models.PowerPointSlide slide, PowerPoint.ShapeRange shapeRange, int percentage)
        {
            if (!IsValidShapeRange(shapeRange))
            {
                return null;
            }

            try
            {
                _slide = slide;

                var hasManyShapes = shapeRange.Count > 1;
                var shape = hasManyShapes ? shapeRange.Group() : shapeRange[1];
                var left = shape.Left;
                var top = shape.Top;
                shapeRange.Cut();

                Utils.Graphics.ExportSlide(_slide, BlurPicture);
                BlurImage(BlurPicture, percentage);

                shapeRange = slide.Shapes.Paste();
                shapeRange.Left = left;
                shapeRange.Top = top;
                if (hasManyShapes)
                {
                    shapeRange = shapeRange.Ungroup();
                }

                var ungroupedRange = UngroupAllShapeRange(shapeRange);
                var shapeGroupNames = ApplyBlurEffect(BlurPicture, ungroupedRange);
                var range = _slide.Shapes.Range(shapeGroupNames.ToArray());

                return range;
            }
            catch (Exception e)
            {
                ActionFramework.Common.Log.Logger.LogException(e, "BlurSelectedEffect");

                ShowErrorMessageBox(e.Message, e);
                return null;
            }
        }

        public static bool IsValidSelection(PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                || selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                return true;
            }

            ShowErrorMessageBox(ErrorMessageNoSelection);
            return false;
        }

        public static bool IsValidShapeRange(PowerPoint.ShapeRange shapeRange)
        {
            if (shapeRange.Count > 0)
            {
                return true;
            }

            ShowErrorMessageBox(ErrorMessageNoSelection);
            return false;
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

        private static List<string> ApplyBlurEffect(string imageFile, PowerPoint.ShapeRange shapeRange)
        {
            var shapeGroupNames = new List<string>();

            for (int i = 0; i < shapeRange.Count; i++)
            {
                var shape = shapeRange[i + 1];

                if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    var shapeNames = ApplyBlurEffectTextBox(imageFile, shape);
                    shapeGroupNames.AddRange(shapeNames);
                }
                else // if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    var shapeName = ApplyBlurEffectShape(imageFile, shape);
                    shapeGroupNames.Add(shapeName);
                }
            }

            return shapeGroupNames;
        }

        private static List<string> ApplyBlurEffectTextBox(string imageFile, PowerPoint.Shape textBox)
        {
            var shapeNames = new List<string>();
            shapeNames.Add(textBox.Name);

            textBox.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            textBox.Fill.Visible = Office.MsoTriState.msoFalse;
            textBox.Line.Visible = Office.MsoTriState.msoFalse;

            var blurShape = _slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, textBox.Left, textBox.Top, textBox.Width,
                        textBox.Height);
            blurShape.Rotation = textBox.Rotation;
            Utils.Graphics.MoveZToJustBehind(blurShape, textBox);
            CropToShape.FillInShapeWithImage(imageFile, blurShape, isInPlace: true);
            shapeNames.Add(blurShape.Name);
            
            if (IsTintSelected)
            {
                var overlayShape = GenerateOverlayShape(_slide, blurShape);
                shapeNames.Add(overlayShape.Name);
            }

            // cannot group placeholders
            if (textBox.Type != Office.MsoShapeType.msoPlaceholder)
            {
                var subRange = _slide.Shapes.Range(shapeNames.ToArray());
                var groupedShape = subRange.Group();
                shapeNames.Clear();
                shapeNames.Add(groupedShape.Name);
            }

            return shapeNames;
        }

        private static string ApplyBlurEffectShape(string imageFile, PowerPoint.Shape shape)
        {
            var shapeNames = new List<string>();
            shapeNames.Add(shape.Name);

            if (!string.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                var textBox = DuplicateShapeInPlace(shape);
                textBox.Fill.Visible = Office.MsoTriState.msoFalse;
                textBox.Line.Visible = Office.MsoTriState.msoFalse;
                Utils.Graphics.MoveZToJustInFront(textBox, shape);
                shapeNames.Add(textBox.Name);
            }

            shape.TextFrame2.DeleteText();
            CropToShape.FillInShapeWithImage(imageFile, shape, isInPlace: true);

            if (IsTintSelected)
            {
                var overlayShape = GenerateOverlayShape(_slide, shape);
                shapeNames.Add(overlayShape.Name);
            }

            if (shapeNames.Count > 1)
            {
                var subRange = _slide.Shapes.Range(shapeNames.ToArray());
                var groupedShape = subRange.Group();

                return groupedShape.Name;
            }

            return shapeNames[0];
        }

        public static void BlurImage(string imageFile, int percentage)
        {
            if (percentage != 0)
            {
                var degree = 50 + (percentage / 2);

                using (var imageFactory = new ImageFactory())
                {
                    var loadedImageFactory = imageFactory.Load(imageFile);
                    var image = loadedImageFactory.Image;
                    var originalWidth = image.Width;
                    var originalHeight = image.Height;

                    var ratio = (float)originalWidth / originalHeight;
                    var targetHeight = Math.Round(1100f - (1100f - 11f) / 100f * degree);
                    var targetWidth = Math.Round(targetHeight * ratio);

                    loadedImageFactory
                        .Resize(new Size((int)targetWidth, (int)targetHeight))
                        .GaussianBlur(5)
                        .Resize(new ResizeLayer(new Size(originalWidth, originalHeight), resizeMode: ResizeMode.Stretch))
                        .Save(imageFile);
                }
            }
        }

        public static PowerPoint.Shape GenerateOverlayShape(Models.PowerPointSlide slide, PowerPoint.Shape blurShape)
        {
            PowerPoint.Shape overlayShape = null;
            
            if (blurShape.Type == Office.MsoShapeType.msoPicture)
            {
                overlayShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, blurShape.Left, blurShape.Top, blurShape.Width,
                    blurShape.Height);
                overlayShape.Rotation = blurShape.Rotation;
            }
            else
            {
                overlayShape = DuplicateShapeInPlace(blurShape);
            }

            Utils.Graphics.MoveZToJustInFront(overlayShape, blurShape);

            var rgb = Utils.Graphics.ConvertColorToRgb(Utils.StringUtil.GetColorFromHexValue(HexColor));

            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = rgb;
            overlayShape.Fill.Transparency = Transparency;
            overlayShape.Line.ForeColor.RGB = rgb;
            overlayShape.Line.Transparency = Transparency;
            overlayShape.Line.Visible = Office.MsoTriState.msoFalse;

            return overlayShape;
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
                || content.Equals(ErrorMessageNoSelection)
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
