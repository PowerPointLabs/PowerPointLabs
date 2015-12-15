using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
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
using Size = System.Drawing.Size;

namespace PowerPointLabs.ImageSearch.Handler
{
    public class EffectsHandler : PowerPointSlide
    {
        private const string ShapeNamePrefix = "pptImagesLab";

        private const float MinThumbnailHeight = 11f;
        private const float MaxThumbnailHeight = 1100f;

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

        public void ApplyImageReferenceInsertion(string contextLink, string fontFamily, string fontColor)
        {
            var imageRefShape = Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, PreviewPresentation.SlideWidth,
                20);
            imageRefShape.TextFrame2.TextRange.Text = "Image From: " + contextLink;

            imageRefShape.TextFrame2.TextRange.TrimText().Font.Fill.ForeColor.RGB 
                = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));
            imageRefShape.TextEffect.FontName = StringUtil.IsEmpty(fontFamily) ? "Tahoma" : fontFamily;
            imageRefShape.TextEffect.FontSize = 14;
            imageRefShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentRight;
            imageRefShape.Top = PreviewPresentation.SlideHeight -
                                imageRefShape.TextFrame2.TextRange.Paragraphs.BoundHeight - 10;
            AddTag(imageRefShape, Tag.ImageReference, "true");
            ChangeName(imageRefShape, EffectName.ImageReference);
        }

        // add a background image shape from imageItem
        public PowerPoint.Shape ApplyBackgroundEffect(string overlayColor, int overlayTransparency, int offset)
        {
            var overlay = ApplyOverlayEffect(overlayColor, overlayTransparency);
            overlay.ZOrder(MsoZOrderCmd.msoSendToBack);

            return ApplyBackgroundEffect(offset);
        }

        public PowerPoint.Shape ApplyBackgroundEffect(int offset)
        {
            var imageShape = AddPicture(Source.FullSizeImageFile ?? Source.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            FitToSlide.AutoFit(imageShape, PreviewPresentation, offset);

            CropPicture(imageShape);
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
                    shape.TextEffect.FontName = shape.Tags[Tag.OriginalFontFamily];
                    shape.Tags.Add(Tag.OriginalFontFamily, "");
                }
                else
                {
                    shape.TextEffect.FontName = fontFamily;
                }

                if (StringUtil.IsEmpty(shape.Tags[Tag.OriginalFontSize]))
                {
                    shape.Tags.Add(Tag.OriginalFontSize, shape.TextEffect.FontSize.ToString(CultureInfo.InvariantCulture));
                }
                shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]) + fontSizeToIncrease;
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

        public void ApplyTextPositionAndAlignment(Position pos, Alignment alignment)
        {
            new TextBoxes(Shapes, PreviewPresentation.SlideWidth, PreviewPresentation.SlideHeight)
                .SetPosition(pos)
                .SetAlignment(alignment)
                .StartBoxing();
        }

        public void ApplyTextWrapping()
        {
            new TextBoxes(Shapes, PreviewPresentation.SlideWidth, PreviewPresentation.SlideHeight)
                .StartTextWrapping();
        }

        // add overlay layer 
        public PowerPoint.Shape ApplyOverlayEffect(string color, int transparency,
            float left = 0, float top = 0, float? width = null, float? height = null)
        {
            width = width ?? PreviewPresentation.SlideWidth;
            height = height ?? PreviewPresentation.SlideHeight;
            var overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top,
                width.Value, height.Value);
            ChangeName(overlayShape, EffectName.Overlay);
            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Fill.Transparency = (float) transparency / 100;
            overlayShape.Line.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Line.Transparency = (float)transparency / 100;
            overlayShape.Line.Weight = 5;
            overlayShape.Line.Visible = MsoTriState.msoFalse;
            return overlayShape;
        }

        public PowerPoint.Shape ApplyCircleOverlayEffect(string color, int transparency,
            float left, float top, float width, float height, bool isOutline)
        {
            var radius = (float) Math.Sqrt(width*width/4 + height*height/4);
            var circleLeft = left - radius + width/2;
            var circleTop = top - radius + height/2;
            var circleWidth = radius*2;

            var overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeOval, circleLeft, circleTop,
                circleWidth, circleWidth);
            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Fill.Transparency = (float)transparency / 100;
            overlayShape.Line.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Line.Transparency = (float)transparency / 100;
            overlayShape.Line.Weight = 5;
            if (isOutline)
            {
                overlayShape.Fill.Visible = MsoTriState.msoFalse;
                overlayShape.Line.Visible = MsoTriState.msoTrue;
            }
            else
            {
                overlayShape.Fill.Visible = MsoTriState.msoTrue;
                overlayShape.Line.Visible = MsoTriState.msoTrue;
                overlayShape.Line.ForeColor.RGB = overlayShape.Fill.ForeColor.RGB;
            }
            // as picture shape
            overlayShape.Cut();
            overlayShape = Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            overlayShape.Left = circleLeft;
            overlayShape.Top = circleTop;
            ChangeName(overlayShape, EffectName.Overlay);
            CropPicture(overlayShape);
            return overlayShape;
        }

        private void CropPicture(PowerPoint.Shape picShape)
        {
            if (picShape.Left < 0)
            {
                picShape.PictureFormat.Crop.ShapeLeft = 0;
            }
            if (picShape.Top < 0)
            {
                picShape.PictureFormat.Crop.ShapeTop = 0;
            }
            if (picShape.Left + picShape.Width > PreviewPresentation.SlideWidth)
            {
                picShape.PictureFormat.Crop.ShapeWidth = PreviewPresentation.SlideWidth - picShape.Left;
            }
            if (picShape.Top + picShape.Height > PreviewPresentation.SlideHeight)
            {
                picShape.PictureFormat.Crop.ShapeHeight = PreviewPresentation.SlideHeight - picShape.Top;
            }
        }

        public PowerPoint.Shape ApplyBlurEffect(string imageFileToBlur = null, int degree = 85, int offset = 0)
        {
            Source.BlurImageFile = BlurImage(imageFileToBlur 
                ?? Source.FullSizeImageFile 
                ?? Source.ImageFile, degree);
            var blurImageShape = AddPicture(Source.BlurImageFile, EffectName.Blur);
            FitToSlide.AutoFit(blurImageShape, PreviewPresentation, offset);
            CropPicture(blurImageShape);
            return blurImageShape;
        }

        public void ApplyTextboxEffect(string overlayColor, int transparency)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.AddedTextbox])
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.ImageReference]))
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
                        var top = paragraph.BoundTop;
                        var width = paragraph.BoundWidth + 10;
                        var height = paragraph.BoundHeight;

                        var overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                            left, top, width, height);
                        ChangeName(overlayShape, EffectName.TextBox);
                        Graphics.MoveZToJustBehind(overlayShape, shape);
                        shape.Tags.Add(Tag.AddedTextbox, overlayShape.Name);
                    }
                }
            }
            foreach (PowerPoint.Shape shape in Shapes)
            {
                shape.Tags.Add(Tag.AddedTextbox, "");
            }
        }

        public PowerPoint.Shape ApplyCircleBannerEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency, 
            bool isOutline = false)
        {
            var tbInfo =
                new TextBoxes(Shapes, PreviewPresentation.SlideWidth, PreviewPresentation.SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;
            TextBoxes.AddMargin(tbInfo, 10);

            var overlayShape = ApplyCircleOverlayEffect(overlayColor, transparency, tbInfo.Left, tbInfo.Top, tbInfo.Width,
                tbInfo.Height, isOutline);
            ChangeName(overlayShape, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return overlayShape;
        }

        public PowerPoint.Shape ApplyCircleOutlineEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            var overlayShape = ApplyCircleBannerEffect(imageShape, overlayColor, transparency, isOutline: true);
            if (overlayShape == null) return null;

            overlayShape.Fill.Visible = MsoTriState.msoFalse;
            overlayShape.Line.Visible = MsoTriState.msoFalse;
            return overlayShape;
        }

        public PowerPoint.Shape ApplyRectBannerEffect(BannerDirection direction, Position textPos, PowerPoint.Shape imageShape, 
            string overlayColor, int transparency)
        {
            var tbInfo =
                new TextBoxes(Shapes, PreviewPresentation.SlideWidth, PreviewPresentation.SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;

            TextBoxes.AddMargin(tbInfo);

            PowerPoint.Shape overlayShape;
            direction = HandleAutoDirection(direction, textPos);
            switch (direction)
            {
                case BannerDirection.Horizontal:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, 0, tbInfo.Top, PreviewPresentation.SlideWidth,
                        tbInfo.Height);
                    break;
                // case BannerDirection.Vertical:
                default:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, 0, tbInfo.Width,
                        PreviewPresentation.SlideHeight);
                    break;
            }
            ChangeName(overlayShape, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return overlayShape;
        }

        public PowerPoint.Shape ApplyRectOutlineEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            var tbInfo =
                new TextBoxes(Shapes, PreviewPresentation.SlideWidth, PreviewPresentation.SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;

            TextBoxes.AddMargin(tbInfo, 10);

            var overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, tbInfo.Top, tbInfo.Width,
                tbInfo.Height);
            overlayShape.Fill.Visible = MsoTriState.msoFalse;
            overlayShape.Line.Visible = MsoTriState.msoTrue;

            ChangeName(overlayShape, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return overlayShape;
        }

        private BannerDirection HandleAutoDirection(BannerDirection dir, Position textPos)
        {
            if (dir != BannerDirection.Auto) return dir;
            switch (textPos)
            {
                case Position.Left:
                case Position.Centre:
                case Position.Right:
                    return BannerDirection.Vertical;
                default:
                    return BannerDirection.Horizontal;
            }
        }

        public PowerPoint.Shape ApplySpecialEffectEffect(IMatrixFilter effectFilter,
            PowerPoint.Shape imageShape, string overlayColor, int transparency, int offset)
        {
            var overlayShape = ApplyOverlayEffect(overlayColor, transparency);
            var specialEffectImageShape = ApplySpecialEffectEffect(effectFilter, false, offset);

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            specialEffectImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return specialEffectImageShape;
        }

        public PowerPoint.Shape ApplySpecialEffectEffect(IMatrixFilter effectFilter, bool isActualSize, int offset)
        {
            Source.SpecialEffectImageFile = SpecialEffectImage(effectFilter, Source.FullSizeImageFile ?? Source.ImageFile, isActualSize);
            var specialEffectImageShape = AddPicture(Source.SpecialEffectImageFile, EffectName.SpecialEffect);
            FitToSlide.AutoFit(specialEffectImageShape, PreviewPresentation, offset);
            CropPicture(specialEffectImageShape);
            return specialEffectImageShape;
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

        public List<PowerPoint.Shape> EmbedStyleOptionsInformation(string originalImageFile, string croppedImageFile, 
            Rect rect, StyleOptions opt)
        {
            if (originalImageFile == null) return new List<PowerPoint.Shape>();

            var originalImage = AddPicture(originalImageFile, EffectName.Original_DO_NOT_REMOVE);
            originalImage.Visible = MsoTriState.msoFalse;

            var croppedImage = AddPicture(croppedImageFile, EffectName.Cropped_DO_NOT_REMOVE);
            croppedImage.Visible = MsoTriState.msoFalse;

            var result = new List<PowerPoint.Shape>();
            result.Add(originalImage);
            result.Add(croppedImage);

            // store source image info
            AddTag(originalImage, Tag.ReloadOriginImg, originalImageFile);
            AddTag(originalImage, Tag.ReloadCroppedImg, croppedImageFile);
            AddTag(originalImage, Tag.ReloadRectX, rect.X.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectY, rect.Y.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectWidth, rect.Width.ToString(CultureInfo.InvariantCulture));
            AddTag(originalImage, Tag.ReloadRectHeight, rect.Height.ToString(CultureInfo.InvariantCulture));

            // store style info
            var type = opt.GetType();
            var props = type.GetProperties();
            foreach (var propertyInfo in props)
            {
                try
                {
                    AddTag(originalImage, Tag.ReloadPrefix + propertyInfo.Name,
                        propertyInfo.GetValue(opt, null).ToString());
                }
                catch (Exception e)
                {
                    PowerPointLabsGlobals.LogException(e, "EmbedStyleOptionsInformation");
                }
            }
            return result;
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

        private static string BlurImage(string imageFilePath, int degree)
        {
            if (degree == 0)
            {
                return imageFilePath;
            }

            var blurImageFile = TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                    .Load(imageFilePath)
                    .Image;
                var ratio = (float) image.Width / image.Height;
                var targetHeight = MaxThumbnailHeight - (MaxThumbnailHeight - MinThumbnailHeight) / 100f * degree;

                image = imageFactory
                    .Resize(new Size((int)(targetHeight * ratio), (int)targetHeight))
                    .GaussianBlur(5).Image;
                image.Save(blurImageFile);
            }
            return blurImageFile;
        }

        public static string SpecialEffectImage(IMatrixFilter effectFilter, string imageFilePath,
            bool isActualSize)
        {
            var specialEffectImageFile = TempPath.GetPath("fullsize_specialeffect");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                        .Load(imageFilePath)
                        .Image;
                var ratio = (float)image.Width / image.Height;
                if (isActualSize)
                {
                    image = imageFactory
                        .Resize(new Size((int) (768 * ratio), 768))
                        .Filter(effectFilter)
                        .Image;
                }
                else
                {
                    image = imageFactory
                        .Resize(new Size((int)(300 * ratio), 300))
                        .Filter(effectFilter)
                        .Image;
                }
                image.Save(specialEffectImageFile);
            }
            return specialEffectImageFile;
        }
        #endregion
    }
}
