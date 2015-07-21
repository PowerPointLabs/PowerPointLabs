using System;
using System.Drawing;
using ImageProcessor;
using Microsoft.Office.Core;
using PowerPointLabs.ImageSearch.Model;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using Graphics = PowerPointLabs.Utils.Graphics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ImageSearch.Slide
{
    public class StylesPreviewSlide : PowerPointSlide
    {
        public enum EffectName
        {
            BackGround,
            Overlay,
            Blur,
            TextBox
        }

        public const string ShapeNamePrefix = "pptImagesLab";

        private ImageItem ImageItem { get; set; }

        private PowerPointPresentation PreviewPresentation { get; set; }

        public StylesPreviewSlide(PowerPoint.Slide slide, PowerPointPresentation pres, ImageItem imageItem)
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

        // add a background image shape from imageItem
        // TODO add image reference info
        public PowerPoint.Shape ApplyBackgroundEffect()
        {
            var imageShape = AddPicture(ImageItem.FullSizeImageFile ?? ImageItem.ImageFile, EffectName.BackGround);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            FitToSlide.AutoFit(imageShape, PreviewPresentation);
            return imageShape;
        }

        // apply text formats to textbox & placeholer
        // TODO think about a way to recover
        public void ApplyTextEffect()
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                shape.Fill.Visible = MsoTriState.msoFalse;
                shape.Line.Visible = MsoTriState.msoFalse;

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                font.Fill.ForeColor.RGB 
                    = Graphics.ConvertColorToRgb(Color.White); // TODO customize
                font.Name = "Segoe UI Light"; // TODO customize
                var textEffect = shape.TextEffect;
//                textEffect.FontSize += 10; // TODO customize
            }
        }

        // add overlay layer 
        public PowerPoint.Shape ApplyOverlayStyle(float transparency)
        {
            var overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0,
                PreviewPresentation.SlideWidth,
                PreviewPresentation.SlideHeight);
            ChangeName(overlayShape, EffectName.Overlay);
            overlayShape.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Black);
            overlayShape.Fill.Transparency = transparency;
            overlayShape.Line.Visible = MsoTriState.msoFalse;
            return overlayShape;
        }

        // add a blured background image shape from imageItem
        public PowerPoint.Shape ApplyBlurEffect(PowerPoint.Shape imageShape = null)
        {
            var overlayShape = ApplyOverlayStyle(0.85f);

            if (ImageItem.BlurImageFile == null)
            {
                ImageItem.BlurImageFile = BlurImage(ImageItem.ImageFile); // TODO customize - can use full-size image
            }
            var blurImageShape = AddPicture(ImageItem.BlurImageFile, EffectName.Blur);
            FitToSlide.AutoFit(blurImageShape, PreviewPresentation);

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return blurImageShape;
        }

        public void ApplyBlurTextboxEffect(PowerPoint.Shape blurImageShape)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse
                        || shape.Tags[ShapeNamePrefix].Trim().Length != 0)
                {
                    continue;
                }

                // multiple paragraphs.. 
                foreach (TextRange2 paragraph in shape.TextFrame2.TextRange.Paragraphs)
                {
                    if (paragraph.TrimText().Length > 0)
                    {
                        blurImageShape.Copy();
                        var blurImageShapeCopy = Shapes.Paste()[1];
                        ChangeName(blurImageShapeCopy, EffectName.Blur);
                        PowerPointLabsGlobals.CopyShapePosition(blurImageShape, ref blurImageShapeCopy);
                        blurImageShapeCopy.PictureFormat.Crop.ShapeLeft = paragraph.BoundLeft - 5;
                        blurImageShapeCopy.PictureFormat.Crop.ShapeWidth = paragraph.BoundWidth + 10;
                        blurImageShapeCopy.PictureFormat.Crop.ShapeTop = paragraph.BoundTop - 5;
                        blurImageShapeCopy.PictureFormat.Crop.ShapeHeight = paragraph.BoundHeight + 10;
                        var overlayBlurShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                            paragraph.BoundLeft - 5,
                            paragraph.BoundTop - 5,
                            paragraph.BoundWidth + 10,
                            paragraph.BoundHeight + 10);
                        overlayBlurShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black);
                        overlayBlurShape.Fill.Transparency = 0.85f;
                        overlayBlurShape.Line.Visible = MsoTriState.msoFalse;
                        ChangeName(overlayBlurShape, EffectName.Blur);
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

        private static string BlurImage(string imageFilePath)
        {
            var blurImageFile = TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory.Load(imageFilePath);
                image = image.GaussianBlur(5);
                image.Save(blurImageFile);
                if (image.MimeType == "image/png")
                {
                    blurImageFile += ".png";
                }
            }
            return blurImageFile;
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
    }
}
