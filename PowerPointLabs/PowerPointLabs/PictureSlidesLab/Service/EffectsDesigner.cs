using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Microsoft.Office.Core;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;
using Graphics = PowerPointLabs.Utils.Graphics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Size = System.Drawing.Size;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    public class EffectsDesigner : PowerPointSlide
    {
        public const string ShapeNamePrefix = "pptPictureSlidesLab";

        private const float MinThumbnailHeight = 11f;
        private const float MaxThumbnailHeight = 1100f;

        private ImageItem Source { get; set; }

        private float SlideWidth { get; set; }

        private float SlideHeight { get; set; }

        private PowerPoint.Slide ContentSlide { get; set; }

        # region APIs
        /// <summary>
        /// For `apply`
        /// </summary>
        /// <param name="slide">the slide to apply the style</param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="source"></param>
        public EffectsDesigner(PowerPoint.Slide slide, float slideWidth, float slideHeight, ImageItem source)
            : base(slide)
        {
            Setup(slideWidth, slideHeight, source);
        }

        /// <summary>
        /// For `preview`
        /// </summary>
        /// <param name="slide">the temp slide to produce preview image</param>
        /// <param name="contentSlide">the slide that contains content</param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="source"></param>
        public EffectsDesigner(PowerPoint.Slide slide, PowerPoint.Slide contentSlide, 
            float slideWidth, float slideHeight, ImageItem source)
            : base(slide)
        {
            ContentSlide = contentSlide;
            Setup(slideWidth, slideHeight, source);
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
            var imageRefShape = Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, SlideWidth,
                20);
            imageRefShape.TextFrame2.TextRange.Text = "Image From: " + contextLink;

            imageRefShape.TextFrame2.TextRange.TrimText().Font.Fill.ForeColor.RGB 
                = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));
            imageRefShape.TextEffect.FontName = StringUtil.IsEmpty(fontFamily) ? "Tahoma" : fontFamily;
            imageRefShape.TextEffect.FontSize = 14;
            imageRefShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentRight;
            imageRefShape.Top = SlideHeight -
                                imageRefShape.TextFrame2.TextRange.Paragraphs.BoundHeight - 10;
            AddTag(imageRefShape, Tag.ImageReference, "true");
            ChangeName(imageRefShape, EffectName.ImageReference);
        }

        public PowerPoint.Shape ApplyBackgroundEffect(int offset)
        {
            var imageShape = AddPicture(Source.FullSizeImageFile ?? Source.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
            FitToSlide.AutoFit(imageShape, slideWidth, slideHeight, offset);

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
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .SetPosition(pos)
                .SetAlignment(alignment)
                .StartBoxing();
        }

        public void ApplyPseudoTextWhenNoTextShapes()
        {
            var isTextShapesEmpty = new TextBoxes(
                Shapes.Range(), SlideWidth, SlideHeight)
                .IsTextShapesEmpty();

            if (!isTextShapesEmpty) return;

            try
            {
                Shapes.AddTitle().TextFrame2.TextRange.Text = "Picture Slides Lab";
            }
            catch
            {
                // title already exist
                foreach (PowerPoint.Shape shape in Shapes)
                {
                    switch (shape.PlaceholderFormat.Type)
                    {
                        case PowerPoint.PpPlaceholderType.ppPlaceholderTitle:
                        case PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle:
                        case PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle:
                            shape.TextFrame2.TextRange.Text = "Picture Slides Lab";
                            break;
                    }
                }
            }
        }

        public void ApplyTextWrapping()
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .StartTextWrapping();
        }

        public void RecoverTextWrapping()
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .RecoverTextWrapping();
        }

        // add overlay layer 

        public PowerPoint.Shape ApplyOverlayEffect(string color, int transparency,
            float left = 0, float top = 0, float? width = null, float? height = null)
        {
            width = width ?? SlideWidth;
            height = height ?? SlideHeight;
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

        public PowerPoint.Shape ApplyBlurEffect(string imageFileToBlur = null, int degree = 85, int offset = 0)
        {
            Source.BlurImageFile = BlurImage(imageFileToBlur 
                ?? Source.FullSizeImageFile 
                ?? Source.ImageFile, degree);
            var blurImageShape = AddPicture(Source.BlurImageFile, EffectName.Blur);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
            FitToSlide.AutoFit(blurImageShape, slideWidth, slideHeight, offset);
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

        public PowerPoint.Shape ApplyCircleBannerEffect(string overlayColor, int transparency, 
            bool isOutline = false, int margin = 0)
        {
            var tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;
            TextBoxes.AddMargin(tbInfo, margin);

            var overlayShape = ApplyCircleOverlayEffect(overlayColor, transparency, tbInfo.Left, tbInfo.Top, tbInfo.Width,
                tbInfo.Height, isOutline);
            ChangeName(overlayShape, EffectName.Banner);
            return overlayShape;
        }

        public PowerPoint.Shape ApplyCircleRingsEffect(string color, int transparency)
        {
            var innerCircleShape = ApplyCircleBannerEffect(color, transparency);
            var outerCircleShape = ApplyCircleBannerEffect(color, transparency, isOutline: true, margin: 10);
            if (innerCircleShape == null || outerCircleShape == null)
            {
                return null;
            }

            outerCircleShape.Left = innerCircleShape.Left + innerCircleShape.Width / 2 - outerCircleShape.Width / 2;
            outerCircleShape.Top = innerCircleShape.Top + innerCircleShape.Height / 2 - outerCircleShape.Height / 2;
            CropPicture(innerCircleShape);
            CropPicture(outerCircleShape);

            var result = Shapes.Range(new[] {innerCircleShape.Name, outerCircleShape.Name}).Group();
            ChangeName(result, EffectName.Overlay);
            return result;
        }

        public PowerPoint.Shape ApplyRectBannerEffect(BannerDirection direction, Position textPos, PowerPoint.Shape imageShape, 
            string overlayColor, int transparency)
        {
            var tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;

            TextBoxes.AddMargin(tbInfo);

            PowerPoint.Shape overlayShape;
            direction = HandleAutoDirection(direction, textPos);
            switch (direction)
            {
                case BannerDirection.Horizontal:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, 0, tbInfo.Top, SlideWidth,
                        tbInfo.Height);
                    break;
                // case BannerDirection.Vertical:
                default:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, 0, tbInfo.Width,
                        SlideHeight);
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
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
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

        public PowerPoint.Shape ApplyAlbumFrameEffect(string overlayColor, int transparency)
        {
            var halfFrameWidth = 15;
            var width = SlideWidth - halfFrameWidth * 2;
            var height = SlideHeight - halfFrameWidth * 2;
            var frameShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, halfFrameWidth, halfFrameWidth,
                width, height);
            ChangeName(frameShape, EffectName.Overlay);
            frameShape.Fill.Transparency = 1f;
            frameShape.Line.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor));
            frameShape.Line.Transparency = (float)transparency / 100;
            frameShape.Line.Weight = 30;
            frameShape.Line.Visible = MsoTriState.msoTrue;
            return frameShape;
        }

        public PowerPoint.Shape ApplyTriangleEffect(string overlayColor1, string overlayColor2, int transparency)
        {
            var width1 = SlideHeight;
            var height1 = SlideWidth;
            var centerLeft1 = SlideWidth/2;
            var centerTop1 = SlideHeight/2;
            // the bigger triangle
            var triangle1 = Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle, 
                centerLeft1 - centerTop1, centerLeft1 + centerTop1 - SlideWidth, width1, height1);
            triangle1.Rotation = 90;
            ChangeName(triangle1, EffectName.Overlay);
            triangle1.Fill.Solid();
            triangle1.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor1));
            triangle1.Fill.Transparency = (float)transparency / 100;
            triangle1.Line.Visible = MsoTriState.msoFalse;

            var width2 = SlideHeight/2;
            var height2 = SlideWidth/2;
            var centerLeft2 = SlideWidth/4*3;
            var centerTop2 = SlideHeight/4*3;
            // the smaller triangle
            var triangle2 = Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle,
                centerLeft2 + centerTop2 - SlideHeight, 
                centerTop2 + SlideWidth/2 - centerLeft2, 
                width2, 
                height2);
            triangle2.Rotation = 270;
            ChangeName(triangle2, EffectName.Overlay);
            triangle2.Fill.Solid();
            triangle2.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor2));
            triangle2.Fill.Transparency = (float)transparency / 100;
            triangle2.Line.Visible = MsoTriState.msoFalse;

            var result = Shapes.Range(new[] {triangle1.Name, triangle2.Name}).Group();
            ChangeName(result, EffectName.Overlay);
            return result;
        }

        public PowerPoint.Shape ApplySpecialEffectEffect(IMatrixFilter effectFilter, bool isActualSize, int offset)
        {
            Source.SpecialEffectImageFile = SpecialEffectImage(effectFilter, Source.FullSizeImageFile ?? Source.ImageFile, isActualSize);
            var specialEffectImageShape = AddPicture(Source.SpecialEffectImageFile, EffectName.SpecialEffect);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
            FitToSlide.AutoFit(specialEffectImageShape, slideWidth, slideHeight, offset);
            CropPicture(specialEffectImageShape);
            return specialEffectImageShape;
        }

        # endregion

        # region Helper Funcs

        private void Setup(float slideWidth, float slideHeight, ImageItem source)
        {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            Source = source;
            PrepareShapesForPreview();
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
            if (picShape.Left + picShape.Width > SlideWidth)
            {
                picShape.PictureFormat.Crop.ShapeWidth = SlideWidth - picShape.Left;
            }
            if (picShape.Top + picShape.Height > SlideHeight)
            {
                picShape.PictureFormat.Crop.ShapeHeight = SlideHeight - picShape.Top;
            }
        }

        private PowerPoint.Shape ApplyCircleOverlayEffect(string color, int transparency,
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
                overlayShape.Line.Visible = MsoTriState.msoFalse;
            }
            // as picture shape
            overlayShape.Cut();
            overlayShape = Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            overlayShape.Left = circleLeft;
            overlayShape.Top = circleTop;
            ChangeName(overlayShape, EffectName.Overlay);
            return overlayShape;
        }

        private void RemovePreviousImageReference()
        {
            NotesPageText = Regex.Replace(NotesPageText, @"^Background image taken from .* on .*\n", "");
        }

        private void PrepareShapesForPreview()
        {
            try
            {
                if (ContentSlide != null && _slide != ContentSlide)
                {
                    // copy shapes from content slide to preview slide
                    DeleteAllShapes();
                    ContentSlide.Shapes.Range().Copy();
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
            string imageContext, Rect rect, StyleOptions opt)
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
            AddTag(originalImage, Tag.ReloadImgContext, imageContext);
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

        /// <summary>
        /// change the shape name, so that they can be managed (eg delete) by name easily
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="effectName"></param>
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
                var targetHeight = Math.Ceiling(MaxThumbnailHeight - (MaxThumbnailHeight - MinThumbnailHeight) / 100f * degree);
                var targetWidth = Math.Ceiling(targetHeight * ratio);

                image = imageFactory
                    .Resize(new Size((int) targetWidth, (int) targetHeight))
                    .GaussianBlur(5).Image;
                image.Save(blurImageFile);
            }
            return blurImageFile;
        }

        private static string SpecialEffectImage(IMatrixFilter effectFilter, string imageFilePath,
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
