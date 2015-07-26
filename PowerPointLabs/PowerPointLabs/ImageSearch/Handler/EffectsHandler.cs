using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Microsoft.Office.Core;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using Graphics = PowerPointLabs.Utils.Graphics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ImageSearch.Handler
{
    public class EffectsHandler : PowerPointSlide
    {
        public enum EffectName
        {
            BackGround,
            Overlay,
            Blur,
            TextBox,
            Grayscale
        }

        public const string ShapeNamePrefix = "pptImagesLab";

        // tag names
        private const string OriginalFontSize = "originalFontSize";
        private const string OriginalFontColor = "originalFontColor";
        private const string OriginalFontFamily = "originalFontFamily";
        private const string OriginalFillVisible = "originalFillVisible";
        private const string OriginalLineVisible = "originalLineVisible";

        private ImageItem ImageItem { get; set; }

        private PowerPointPresentation PreviewPresentation { get; set; }

        public EffectsHandler(PowerPoint.Slide slide, PowerPointPresentation pres, ImageItem imageItem)
            : base(slide)
        {
            PreviewPresentation = pres;
            ImageItem = imageItem;
            PrepareShapesForPreview();
        }

        private void PrepareShapesForPreview()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                if (_slide != currentSlide.GetNativeSlide())
                {
                    DeleteAllShapes();
                    currentSlide.Shapes.Range().Copy();
                    _slide.Shapes.Paste();
                }
                RemoveAnyStyles();
            }
            catch
            {
                // nothing to copy-paste
            }
        }

        public void RemoveAnyStyles()
        {
            // cannot restore text format though..
            DeleteShapesWithPrefix(ShapeNamePrefix);
        }

        public void RemoveStyles(EffectName effectName)
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

        private void RemovePreviousImageReference()
        {
            NotesPageText = Regex.Replace(NotesPageText, @"^Background image taken from .* on .*\n", "");
        }

        // add a background image shape from imageItem
        public PowerPoint.Shape ApplyBackgroundEffect(string overlayColor,int overlayTransparency)
        {
            var overlay = ApplyOverlayStyle(overlayColor, overlayTransparency);
            overlay.ZOrder(MsoZOrderCmd.msoSendToBack);

            return ApplyBackgroundEffect();
        }

        public PowerPoint.Shape ApplyBackgroundEffect()
        {
            var imageShape = AddPicture(ImageItem.FullSizeImageFile ?? ImageItem.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            FitToSlide.AutoFit(imageShape, PreviewPresentation);

            return imageShape;
        }

        // TODO shorten this?
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

                if (IsEmpty(shape.Tags[OriginalFillVisible]))
                {
                    shape.Tags.Add(OriginalFillVisible, ToBool(shape.Fill.Visible).ToString());
                }
                shape.Fill.Visible = MsoTriState.msoFalse;

                if (IsEmpty(shape.Tags[OriginalLineVisible]))
                {
                    shape.Tags.Add(OriginalLineVisible, ToBool(shape.Line.Visible).ToString());
                }
                shape.Line.Visible = MsoTriState.msoFalse;

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                if (IsEmpty(shape.Tags[OriginalFontColor]))
                {
                    shape.Tags.Add(OriginalFontColor, GetHexValue(Graphics.ConvertRgbToColor(font.Fill.ForeColor.RGB)));
                }
                font.Fill.ForeColor.RGB
                        = Graphics.ConvertColorToRgb(ColorTranslator.FromHtml(fontColor));

                if (IsEmpty(shape.Tags[OriginalFontFamily]))
                {
                    shape.Tags.Add(OriginalFontFamily, font.Name);
                }
                font.Name = IsEmpty(fontFamily) ? shape.Tags[OriginalFontFamily] : fontFamily;

                if (IsEmpty(shape.Tags[OriginalFontSize]))
                {
                    shape.Tags.Add(OriginalFontSize, shape.TextEffect.FontSize.ToString(CultureInfo.InvariantCulture));
                    shape.TextEffect.FontSize += fontSizeToIncrease;
                }
                else // applied before
                {
                    shape.TextEffect.FontSize = float.Parse(shape.Tags[OriginalFontSize]);
                    shape.TextEffect.FontSize += fontSizeToIncrease;
                }
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

                if (!IsEmpty(shape.Tags[OriginalFillVisible]))
                {
                    shape.Fill.Visible = ToMsoTriState(bool.Parse(shape.Tags[OriginalFillVisible]));
                }
                if (!IsEmpty(shape.Tags[OriginalLineVisible]))
                {
                    shape.Line.Visible = ToMsoTriState(bool.Parse(shape.Tags[OriginalLineVisible]));
                }

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                if (!IsEmpty(shape.Tags[OriginalFontColor]))
                {
                    font.Fill.ForeColor.RGB
                        = Graphics.ConvertColorToRgb(ColorTranslator.FromHtml(shape.Tags[OriginalFontColor]));
                }
                if (!IsEmpty(shape.Tags[OriginalFontFamily]))
                {
                    font.Name = shape.Tags[OriginalFontFamily];
                }
                if (!IsEmpty(shape.Tags[OriginalFontSize]))
                {
                    shape.TextEffect.FontSize = float.Parse(shape.Tags[OriginalFontSize]);
                }
            }
        }

        // TODO util
        private bool IsEmpty(string str)
        {
            return str.Trim().Length == 0;
        }

        private bool ToBool(MsoTriState state)
        {
            return state == MsoTriState.msoTrue;
        }

        private MsoTriState ToMsoTriState(bool boolean)
        {
            return boolean ? MsoTriState.msoTrue : MsoTriState.msoFalse;
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
            overlayShape.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(ColorTranslator.FromHtml(color));
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
            if (ImageItem.BlurImageFile == null)
            {
                ImageItem.BlurImageFile = BlurImage(ImageItem.ImageFile);
            }
            var blurImageShape = AddPicture(ImageItem.BlurImageFile, EffectName.Blur);
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
                        || !IsEmpty(shape.Tags[ShapeNamePrefix]))
                {
                    continue;
                }

                // multiple paragraphs.. 
                foreach (TextRange2 paragraph in shape.TextFrame2.TextRange.Paragraphs)
                {
                    if (paragraph.TrimText().Length > 0)
                    {
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
                        shape.Tags.Add(ShapeNamePrefix, blurImageShapeCopy.Name);
                    }
                }
            }
            foreach (PowerPoint.Shape shape in Shapes)
            {
                shape.Tags.Add(ShapeNamePrefix, "");
            }
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        public PowerPoint.Shape ApplyGrayscaleEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            var overlayShape = ApplyOverlayStyle(overlayColor, transparency);

            if (ImageItem.GrayscaleImageFile == null && ImageItem.FullSizeImageFile == null)
            {
                ImageItem.GrayscaleImageFile = GrayscaleImage(ImageItem.ImageFile);
            }
            if (ImageItem.FullSizeGrayscaleImageFile == null && ImageItem.FullSizeImageFile != null)
            {
                ImageItem.FullSizeGrayscaleImageFile = GrayscaleImage(ImageItem.FullSizeImageFile);
                ImageItem.GrayscaleImageFile = ImageItem.FullSizeGrayscaleImageFile;
            }
            var grayscaleImageShape = AddPicture(ImageItem.GrayscaleImageFile, EffectName.Grayscale);
            FitToSlide.AutoFit(grayscaleImageShape, PreviewPresentation);

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            grayscaleImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return grayscaleImageShape;
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

        private string BlurImage(string imageFilePath)
        {
            var blurImageFile = TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                if (ImageItem.FullSizeImageFile == imageFilePath)
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

        private string GrayscaleImage(string imageFilePath)
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
            shape.Name = ShapeNamePrefix + "_" + effectName + "_" + Guid.NewGuid().ToString().Substring(0, 7);
        }

        // TODO util
        private string GetHexValue(Color color)
        {
            byte[] rgbArray = { color.R, color.G, color.B };
            var hex = BitConverter.ToString(rgbArray);
            return "#" + hex.Replace("-", "");
        }
    }
}
