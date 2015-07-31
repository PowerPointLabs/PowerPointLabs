using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Microsoft.Office.Core;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler.Effect;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using Graphics = PowerPointLabs.Utils.Graphics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ImageSearch.Handler
{
    public class EffectsHandler : PowerPointSlide
    {
        private const string ShapeNamePrefix = "pptImagesLab";

        private ImageItem Source { get; set; }

        private PowerPointPresentation PreviewPresentation { get; set; }

        # region APIs
        public EffectsHandler(PowerPoint.Slide slide, PowerPointPresentation pres, ImageItem source)
            : base(slide)
        {
            PreviewPresentation = pres;
            Source = source;
            PrepareShapesForPreview();
        }

        public void RemoveEffect(EffectName effectName)
        {
            DeleteShapesWithPrefix(ShapeNamePrefix + "_" + effectName);
        }

        public void ApplyImageReference(string contextLink)
        {
            if (StringUtil.IsEmpty(contextLink)) return;

            RemovePreviousImageReference();
            NotesPageText = "Background image taken from " + contextLink + " on " + DateTime.Now + "\n" +
                            NotesPageText;
        }

        // add a background image shape from imageItem
        public PowerPoint.Shape ApplyBackgroundEffect(string overlayColor, int overlayTransparency)
        {
            var overlay = ApplyOverlayStyle(overlayColor, overlayTransparency);
            overlay.ZOrder(MsoZOrderCmd.msoSendToBack);

            return ApplyBackgroundEffect();
        }

        public PowerPoint.Shape ApplyBackgroundEffect()
        {
            var imageShape = AddPicture(Source.FullSizeImageFile ?? Source.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            FitToSlide.AutoFit(imageShape, PreviewPresentation);

            return imageShape;
        }

        // apply text formats to textbox & placeholer
        public void ApplyTextEffect(string fontFamily, string fontColor, int fontSizeToIncrease)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                AddTag(shape, Tag.OriginalFillVisible, BoolUtil.ToBool(shape.Fill.Visible).ToString());
                shape.Fill.Visible = MsoTriState.msoFalse;

                AddTag(shape, Tag.OriginalLineVisible, BoolUtil.ToBool(shape.Line.Visible).ToString());
                shape.Line.Visible = MsoTriState.msoFalse;

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                AddTag(shape, Tag.OriginalFontColor, StringUtil.GetHexValue(Graphics.ConvertRgbToColor(font.Fill.ForeColor.RGB)));
                font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));

                AddTag(shape, Tag.OriginalFontFamily, font.Name);
                if (StringUtil.IsEmpty(fontFamily))
                {
                    font.Name = shape.Tags[Tag.OriginalFontFamily];
                    shape.Tags.Add(Tag.OriginalFontFamily, "");
                }
                else
                {
                    font.Name = fontFamily;
                }

                if (StringUtil.IsEmpty(shape.Tags[Tag.OriginalFontSize]))
                {
                    shape.Tags.Add(Tag.OriginalFontSize, shape.TextEffect.FontSize.ToString(CultureInfo.InvariantCulture));
                }
                else // applied before
                {
                    shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]);
                }
                shape.TextEffect.FontSize += fontSizeToIncrease;
            }
        }

        public void ApplyOriginalTextEffect()
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFillVisible]))
                {
                    shape.Fill.Visible = BoolUtil.ToMsoTriState(bool.Parse(shape.Tags[Tag.OriginalFillVisible]));
                    shape.Tags.Add(Tag.OriginalFillVisible, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalLineVisible]))
                {
                    shape.Line.Visible = BoolUtil.ToMsoTriState(bool.Parse(shape.Tags[Tag.OriginalLineVisible]));
                    shape.Tags.Add(Tag.OriginalLineVisible, "");
                }

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontColor]))
                {
                    font.Fill.ForeColor.RGB
                        = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(shape.Tags[Tag.OriginalFontColor]));
                    shape.Tags.Add(Tag.OriginalFontColor, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontFamily]))
                {
                    font.Name = shape.Tags[Tag.OriginalFontFamily];
                    shape.Tags.Add(Tag.OriginalFontFamily, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontSize]))
                {
                    shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]);
                    shape.Tags.Add(Tag.OriginalFontSize, "");
                }
            }
        }

        // add overlay layer 
        public PowerPoint.Shape ApplyOverlayStyle(string color, int transparency,
            float left = 0, float top = 0, float? width = null, float? height = null)
        {
            width = width ?? PreviewPresentation.SlideWidth;
            height = height ?? PreviewPresentation.SlideHeight;
            var overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top,
                width.Value, height.Value);
            ChangeName(overlayShape, EffectName.Overlay);
            overlayShape.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Fill.Transparency = (float) transparency / 100;
            overlayShape.Line.Visible = MsoTriState.msoFalse;
            return overlayShape;
        }

        // add a blured background image shape from imageItem
        public PowerPoint.Shape ApplyBlurEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            var overlayShape = ApplyOverlayStyle(overlayColor, transparency);
            var blurImageShape = ApplyBlurEffect();

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return blurImageShape;
        }

        public PowerPoint.Shape ApplyBlurEffect()
        {
            if (Source.BlurImageFile == null)
            {
                Source.BlurImageFile = BlurImage(Source.ImageFile, 
                    Source.ImageFile == Source.FullSizeImageFile);
            }
            var blurImageShape = AddPicture(Source.BlurImageFile, EffectName.Blur);
            FitToSlide.AutoFit(blurImageShape, PreviewPresentation);
            return blurImageShape;
        }

        public void ApplyBlurTextboxEffect(PowerPoint.Shape blurImageShape, string overlayColor, int transparency)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.AddedTextbox]))
                {
                    continue;
                }

                // multiple paragraphs.. 
                foreach (TextRange2 textRange in shape.TextFrame2.TextRange.Paragraphs)
                {
                    if (StringUtil.IsNotEmpty(textRange.TrimText().Text))
                    {
                        var paragraph = textRange.TrimText();
                        var left = paragraph.BoundLeft - 5;
                        var top = paragraph.BoundTop - 5;
                        var width = paragraph.BoundWidth + 10;
                        var height = paragraph.BoundHeight + 10;

                        var blurImageShapeCopy = BlurTextbox(blurImageShape,
                            left, top, width, height);
                        var overlayBlurShape = ApplyOverlayStyle(overlayColor, transparency,
                            left, top, width, height);
                        Graphics.MoveZToJustBehind(blurImageShapeCopy, shape);
                        Graphics.MoveZToJustBehind(overlayBlurShape, shape);
                        shape.Tags.Add(Tag.AddedTextbox, blurImageShapeCopy.Name);
                    }
                }
            }
            foreach (PowerPoint.Shape shape in Shapes)
            {
                shape.Tags.Add(Tag.AddedTextbox, "");
            }
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        public PowerPoint.Shape ApplyGrayscaleEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            var overlayShape = ApplyOverlayStyle(overlayColor, transparency);

            if (Source.GrayscaleImageFile == null && Source.FullSizeImageFile == null)
            {
                Source.GrayscaleImageFile = GrayscaleImage(Source.ImageFile);
            }
            if (Source.FullSizeGrayscaleImageFile == null && Source.FullSizeImageFile != null)
            {
                Source.FullSizeGrayscaleImageFile = GrayscaleImage(Source.FullSizeImageFile);
                Source.GrayscaleImageFile = Source.FullSizeGrayscaleImageFile;
            }
            var grayscaleImageShape = AddPicture(Source.GrayscaleImageFile, EffectName.Grayscale);
            FitToSlide.AutoFit(grayscaleImageShape, PreviewPresentation);

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            grayscaleImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return grayscaleImageShape;
        }

        # endregion

        # region Helper Funcs
        private void RemovePreviousImageReference()
        {
            NotesPageText = Regex.Replace(NotesPageText, @"^Background image taken from .* on .*\n", "");
        }

        private void PrepareShapesForPreview()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                if (currentSlide != null && _slide != currentSlide.GetNativeSlide())
                {
                    DeleteAllShapes();
                    currentSlide.Shapes.Range().Copy();
                    _slide.Shapes.Paste();
                }
                DeleteShapesWithPrefix(ShapeNamePrefix);
            }
            catch
            {
                // nothing to copy-paste
            }
        }

        private PowerPoint.Shape BlurTextbox(PowerPoint.Shape blurImageShape, 
            float left, float top, float width, float height)
        {
            blurImageShape.Copy();
            var blurImageShapeCopy = Shapes.Paste()[1];
            ChangeName(blurImageShapeCopy, EffectName.Blur);
            PowerPointLabsGlobals.CopyShapePosition(blurImageShape, ref blurImageShapeCopy);
            blurImageShapeCopy.PictureFormat.Crop.ShapeLeft = left;
            blurImageShapeCopy.PictureFormat.Crop.ShapeTop = top;
            blurImageShapeCopy.PictureFormat.Crop.ShapeWidth = width;
            blurImageShapeCopy.PictureFormat.Crop.ShapeHeight = height;
            return blurImageShapeCopy;
        }

        private PowerPoint.Shape AddPicture(string imageFile, EffectName effectName)
        {
            var imageShape = Shapes.AddPicture(imageFile,
                MsoTriState.msoFalse, MsoTriState.msoTrue, 0,
                0);
            ChangeName(imageShape, effectName);
            return imageShape;
        }

        private static void ChangeName(PowerPoint.Shape shape, EffectName effectName)
        {
            ShapeUtil.ChangeName(shape, effectName, ShapeNamePrefix);
        }

        private static void AddTag(PowerPoint.Shape shape, string tagName, String value)
        {
            ShapeUtil.AddTag(shape, tagName, value);
        }

        public static string BlurImage(string imageFilePath, bool isBlurForFullsize)
        {
            var blurImageFile = TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                if (isBlurForFullsize)
                {// for full-size image, need to resize first
                    var image = imageFactory
                        .Load(imageFilePath)
                        .Image;
                    image = imageFactory
                        .Resize(new Size(image.Width / 4, image.Height / 4))
                        .GaussianBlur(5).Image;
                    image.Save(blurImageFile);
                }
                else
                {
                    var image = imageFactory
                        .Load(imageFilePath)
                        .GaussianBlur(5)
                        .Image;
                    image.Save(blurImageFile);
                }
            }
            return blurImageFile;
        }

        public static string GrayscaleImage(string imageFilePath)
        {
            var grayscaleImageFile = TempPath.GetPath("fullsize_grayscale");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                    .Load(imageFilePath)
                    .Filter(MatrixFilters.GreyScale)
                    .Image;
                image.Save(grayscaleImageFile);
            }
            return grayscaleImageFile;
        }
        #endregion
    }
}
