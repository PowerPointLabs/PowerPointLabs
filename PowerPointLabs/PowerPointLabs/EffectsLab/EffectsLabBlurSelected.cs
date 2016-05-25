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

        public static void BlurSelected(Models.PowerPointSlide slide, PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes
                && selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            BlurSelected(slide, selection.ShapeRange);
        }

        public static void BlurSelected(Models.PowerPointSlide slide, PowerPoint.ShapeRange shapeRange)
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
                var shapeGroups = ApplyBlurEffect(BlurPicture, ungroupedRange);
                GroupAndReorderShapes(shapeGroups);
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

        private static List<PowerPoint.Shape>[] ApplyBlurEffect(string imageFile, PowerPoint.ShapeRange shapeRange)
        {
            var shapeGroups = new List<PowerPoint.Shape>[shapeRange.Count];

            for (int i = 0; i < shapeRange.Count; i++)
            {
                var shape = shapeRange[i + 1];
                var group = new List<PowerPoint.Shape>();

                if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    var blurShape = _slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, shape.Left, shape.Top, shape.Width,
                        shape.Height);
                    blurShape.Rotation = shape.Rotation;
                    CropToShape.FillInShapeWithImage(imageFile, blurShape, isInPlace: true);
                    group.Add(blurShape);

                    shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                    group.Add(shape);
                }
                else // if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    group.Add(shape);

                    if (!string.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
                    {
                        var textBox = DuplicateShapeInPlace(shape);
                        textBox.Fill.Visible = Office.MsoTriState.msoFalse;
                        textBox.Line.Visible = Office.MsoTriState.msoFalse;
                        group.Add(textBox);
                    }

                    shape.TextFrame2.DeleteText();
                    CropToShape.FillInShapeWithImage(imageFile, shape, isInPlace: true);
                }

                shapeGroups[i] = group;
            }

            return shapeGroups;
        }

        private static void GroupAndReorderShapes(List<PowerPoint.Shape>[] shapeGroups)
        {
            var shapeGroupNames = new string[shapeGroups.Length];

            for (int i = 0; i < shapeGroups.Length; i++)
            {
                var group = shapeGroups[i];

                if (group.Count == 2)
                {
                    group[1].ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                    var subRange = _slide.Shapes.Range(new[] { group[0].Name, group[1].Name });
                    var groupedShape = subRange.Group();
                    groupedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    shapeGroupNames[i] = groupedShape.Name;
                }
                else
                {
                    shapeGroupNames[i] = group[0].Name;
                }
            }

            var range = _slide.Shapes.Range(shapeGroupNames);
            range.Group();
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
