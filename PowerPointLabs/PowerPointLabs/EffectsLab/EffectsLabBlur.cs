using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

using ImageProcessor;
using ImageProcessor.Imaging;

using PowerPointLabs.CropLab;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabBlur
    {
        private const string HexColor = "#000000";
        private const float Transparency = 0.8f;

        private static readonly string BlurPicture = Path.GetTempPath() + @"\blur.png";

        public static PowerPoint.ShapeRange ExecuteBlurSelected(Models.PowerPointSlide slide, PowerPoint.Selection selection, int percentage)
        {
            if (!IsValidSelection(selection))
            {
                return null;
            }

            PowerPoint.ShapeRange range = BlurSelected(slide, selection, percentage);
            if (range != null)
            {
                range.Select();
            }

            return range;
        }

        public static PowerPoint.ShapeRange BlurSelected(Models.PowerPointSlide slide, PowerPoint.Selection selection, int percentage)
        {
            PowerPoint.ShapeRange shapeRange = ShapeUtil.GetShapeRange(selection);

            try
            {
                bool hasManyShapes = shapeRange.Count > 1;
                PowerPoint.Shape shape = hasManyShapes ? shapeRange.Group() : shapeRange[1];
                float left = shape.Left;
                float top = shape.Top;
                shapeRange.Cut();

                Utils.GraphicsUtil.ExportSlide(slide, BlurPicture);
                BlurImage(BlurPicture, percentage);

                shapeRange = slide.Shapes.Paste();
                shapeRange.Left = left;
                shapeRange.Top = top;
                if (hasManyShapes)
                {
                    shapeRange = shapeRange.Ungroup();
                }

                PowerPoint.ShapeRange ungroupedRange = EffectsLabUtil.UngroupAllShapeRange(slide, shapeRange);
                List<string> shapeGroupNames = ApplyBlurEffect(slide, BlurPicture, ungroupedRange);
                PowerPoint.ShapeRange range = slide.Shapes.Range(shapeGroupNames.ToArray());

                return range;
            }
            catch (Exception e)
            {
                ActionFramework.Common.Log.Logger.LogException(e, "BlurSelectedEffect");

                EffectsLabUtil.ShowErrorMessageBox(e.Message, e);
                return null;
            }
        }

        public static void ExecuteBlurRemainder(Models.PowerPointSlide slide, PowerPoint.Selection selection, int percentage)
        {
            Models.PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(slide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlurBackground(percentage, EffectsLabSettings.IsTintRemainder);
            effectSlide.GetNativeSlide().Select();
        }

        public static void ExecuteBlurBackground(Models.PowerPointSlide slide, PowerPoint.Selection selection, int percentage)
        {
            Models.PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(slide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlurBackground(percentage, EffectsLabSettings.IsTintBackground);
            effectSlide.GetNativeSlide().Select();
        }


        public static void BlurImage(string imageFile, int percentage)
        {
            if (percentage != 0)
            {
                int degree = 50 + (percentage / 2);

                using (ImageFactory imageFactory = new ImageFactory())
                {
                    ImageFactory loadedImageFactory = imageFactory.Load(imageFile);
                    Image image = loadedImageFactory.Image;
                    int originalWidth = image.Width;
                    int originalHeight = image.Height;

                    float ratio = (float)originalWidth / originalHeight;
                    double targetHeight = Math.Round(1100f - (1100f - 11f) / 100f * degree);
                    double targetWidth = Math.Round(targetHeight * ratio);

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
                overlayShape = EffectsLabUtil.DuplicateShapeInPlace(blurShape);
            }

            Utils.ShapeUtil.MoveZToJustInFront(overlayShape, blurShape);

            int rgb = Utils.GraphicsUtil.ConvertColorToRgb(Utils.StringUtil.GetColorFromHexValue(HexColor));

            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = rgb;
            overlayShape.Fill.Transparency = Transparency;
            overlayShape.Line.ForeColor.RGB = rgb;
            overlayShape.Line.Transparency = Transparency;
            overlayShape.Line.Visible = Office.MsoTriState.msoFalse;

            return overlayShape;
        }

        public static bool IsValidSelection(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return IsValidShapeRange(selection.ChildShapeRange);
            }

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                return IsValidShapeRange(selection.ShapeRange);
            }

            EffectsLabUtil.ShowNoSelectionErrorMessage();
            return false;
        }

        public static bool IsValidShapeRange(PowerPoint.ShapeRange shapeRange)
        {
            if (shapeRange.Count > 0)
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    if (shapeRange[i].Type != Office.MsoShapeType.msoPlaceholder &&
                        shapeRange[i].Type != Office.MsoShapeType.msoTextBox &&
                        shapeRange[i].Type != Office.MsoShapeType.msoAutoShape &&
                        shapeRange[i].Type != Office.MsoShapeType.msoFreeform &&
                        shapeRange[i].Type != Office.MsoShapeType.msoGroup)
                    {
                        EffectsLabUtil.ShowIncorrectSelectionErrorMessage();
                        return false;
                    }
                }
            }
            else
            {
                EffectsLabUtil.ShowNoSelectionErrorMessage();
                return false;
            }
            return true;
        }

        private static List<string> ApplyBlurEffect(Models.PowerPointSlide slide, string imageFile, PowerPoint.ShapeRange shapeRange)
        {
            List<string> shapeGroupNames = new List<string>();

            for (int i = 0; i < shapeRange.Count; i++)
            {
                PowerPoint.Shape shape = shapeRange[i + 1];

                if (shape.Type == Office.MsoShapeType.msoPlaceholder
                    || shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    List<string> shapeNames = ApplyBlurEffectTextBox(slide, imageFile, shape);
                    shapeGroupNames.AddRange(shapeNames);
                }
                else // if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    string shapeName = ApplyBlurEffectShape(slide, imageFile, shape);
                    shapeGroupNames.Add(shapeName);
                }
            }

            return shapeGroupNames;
        }

        private static List<string> ApplyBlurEffectTextBox(Models.PowerPointSlide slide, string imageFile, PowerPoint.Shape textBox)
        {
            List<string> shapeNames = new List<string>();
            shapeNames.Add(textBox.Name);

            textBox.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            textBox.Fill.Visible = Office.MsoTriState.msoFalse;
            textBox.Line.Visible = Office.MsoTriState.msoFalse;

            PowerPoint.Shape blurShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, textBox.Left, textBox.Top, textBox.Width,
                        textBox.Height);
            blurShape.Rotation = textBox.Rotation;
            Utils.ShapeUtil.MoveZToJustBehind(blurShape, textBox);
            CropToShape.FillInShapeWithImage(slide, imageFile, blurShape, isInPlace: true);
            shapeNames.Add(blurShape.Name);
            
            if (EffectsLabSettings.IsTintSelected)
            {
                PowerPoint.Shape overlayShape = GenerateOverlayShape(slide, blurShape);
                shapeNames.Add(overlayShape.Name);
            }

            // cannot group placeholders
            if (textBox.Type != Office.MsoShapeType.msoPlaceholder)
            {
                PowerPoint.ShapeRange subRange = slide.Shapes.Range(shapeNames.ToArray());
                PowerPoint.Shape groupedShape = subRange.Group();
                shapeNames.Clear();
                shapeNames.Add(groupedShape.Name);
            }

            return shapeNames;
        }

        private static string ApplyBlurEffectShape(Models.PowerPointSlide slide, string imageFile, PowerPoint.Shape shape)
        {
            List<string> shapeNames = new List<string>();
            shapeNames.Add(shape.Name);

            if (!string.IsNullOrWhiteSpace(shape.TextFrame2.TextRange.Text))
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                PowerPoint.Shape textBox = EffectsLabUtil.DuplicateShapeInPlace(shape);
                textBox.Fill.Visible = Office.MsoTriState.msoFalse;
                textBox.Line.Visible = Office.MsoTriState.msoFalse;
                Utils.ShapeUtil.MoveZToJustInFront(textBox, shape);
                shapeNames.Add(textBox.Name);
            }

            shape.TextFrame2.DeleteText();
            CropToShape.FillInShapeWithImage(slide, imageFile, shape, isInPlace: true);

            if (EffectsLabSettings.IsTintSelected)
            {
                PowerPoint.Shape overlayShape = GenerateOverlayShape(slide, shape);
                shapeNames.Add(overlayShape.Name);
            }

            if (shapeNames.Count > 1)
            {
                PowerPoint.ShapeRange subRange = slide.Shapes.Range(shapeNames.ToArray());
                PowerPoint.Shape groupedShape = subRange.Group();

                return groupedShape.Name;
            }

            return shapeNames[0];
        }
    }
}
