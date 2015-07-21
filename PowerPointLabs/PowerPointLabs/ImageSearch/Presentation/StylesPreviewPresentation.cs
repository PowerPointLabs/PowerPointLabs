using System.Collections.Generic;
using System.Drawing;
using ImageProcessor;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using Microsoft.Office.Core;
using PowerPointLabs.ImageSearch.Model;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ImageSearch.Presentation
{
    public class StylesPreviewPresentation : PowerPointPresentation
    {
        public string TextboxStyleImagePath { get; private set; }
        public string BlurStyleImagePath { get; private set; }
        public string DirectTextStyleImagePath { get; private set; }

        public StylesPreviewPresentation()
        {
            Path = TempPath.TempFolder;
            Name = "ImagesLabPreview";
        }

        public PowerPointSlide AddSlide(PpSlideLayout layout = PpSlideLayout.ppLayoutText)
        {
            if (!Opened)
            {
                return null;
            }

            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);
            var slideFromFactory = PowerPointSlide.FromSlideFactory(newSlide);

            Slides.Add(slideFromFactory);

            return slideFromFactory;
        }

        public void PreviewStyles(ImageItem imageItem)
        {
            InitImagePath();
            
            var thisSlide = AddSlide(PowerPointCurrentPresentationInfo.CurrentSlide.Layout);
            try
            {
                thisSlide.DeleteAllShapes();

                var range = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range();
                thisSlide.CopyShapesToSlide(range);
                thisSlide.DeleteShapesWithPrefix("pptImagesLab");
            }
            catch
            {
                // nothing to copy-paste
            }

            var imageShape = thisSlide.Shapes.AddPicture(imageItem.FullSizeImageFile ?? imageItem.ImageFile,
                MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
            FitToSlide.AutoFit(imageShape, this);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);

            // Textbox style 1 starts
            foreach (PowerPoint.Shape shape in thisSlide.Shapes)
            {
                if (shape.Type == MsoShapeType.msoPlaceholder
                    || shape.Type == MsoShapeType.msoTextBox)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoFalse
                        || shape.Tags["GotHighlighted"].Trim().Length != 0)
                    {
                        continue;
                    }

                    // filled by added shape (can control size)
                    shape.Fill.Visible = MsoTriState.msoFalse;
                    shape.Line.Visible = MsoTriState.msoFalse;

                    var whiteColor = Color.White;// TODO customize
                    var r = whiteColor.R;
                    var g = whiteColor.G;
                    var b = whiteColor.B;

                    var rgb = (b << 16) | (g << 8) | (r);
                    var font = shape.TextFrame2.TextRange.TrimText().Font;
                    font.Fill.ForeColor.RGB = rgb;
                    font.Bold = MsoTriState.msoFalse;
                    font.Name = "Segoe UI Light"; // TODO customize

                    var textEffect = shape.TextEffect;
                    //                        textEffect.FontSize += 10; // TODO customize
                }
            }
            thisSlide.GetNativeSlide().Export(DirectTextStyleImagePath, "JPG");
            
            // Textbox style 1 ends
            //                // Textbox style 2 starts
            //                var overlayShape = thisSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
            //                                                                0,
            //                                                                0,
            //                                                                PowerPointPresentation.Current.SlideWidth,
            //                                                                PowerPointPresentation.Current.SlideHeight);
            //                overlayShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black);
            //                overlayShape.Fill.Transparency = 0.65f;
            //                overlayShape.Line.Visible = MsoTriState.msoFalse;
            //                overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            //                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            //                thisSlide.GetNativeSlide().Export(previewFile3, "JPG");
            //                PreviewList.Add(new ImageItem
            //                {
            //                    ImageFile = previewFile3
            //                });
            //                overlayShape.Delete();
            //
            // textbox style 5 starts
            var overlayShape = thisSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                            0,
                                                            0,
                                                            PowerPointPresentation.Current.SlideWidth,
                                                            PowerPointPresentation.Current.SlideHeight);
            overlayShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black);
            overlayShape.Fill.Transparency = 0.85f;
            overlayShape.Line.Visible = MsoTriState.msoFalse;

            if (imageItem.BlurImageFile == null)
            {
                var blurImageFile = TempPath.GetPath("fullsize_blur");
                using (var imageFactory = new ImageFactory())
                {
                    var image = imageFactory.Load(imageItem.ImageFile);
                    image = image.GaussianBlur(5);
                    image.Save(blurImageFile);
                    if (image.MimeType == "image/png")
                    {
                        blurImageFile += ".png";
                    }
                    imageItem.BlurImageFile = blurImageFile;
                }
            }
            var blurImageShape = thisSlide.Shapes.AddPicture(imageItem.BlurImageFile,
                MsoTriState.msoFalse, MsoTriState.msoTrue, 0,
                0);
            FitToSlide.AutoFit(blurImageShape, this);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);

            thisSlide.GetNativeSlide().Export(BlurStyleImagePath, "JPG");

            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);

            // blur textbox region starts
            var listOfBlurImageCopy = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in thisSlide.Shapes)
            {
                if (shape.Type == MsoShapeType.msoPlaceholder
                    || shape.Type == MsoShapeType.msoTextBox)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoFalse
                        || shape.Tags["GotBlured"].Trim().Length != 0)
                    {
                        continue;
                    }
                    // multiple paragraphs.. 
                    foreach (TextRange2 paragraph in shape.TextFrame2.TextRange.Paragraphs)
                    {
                        if (paragraph.TrimText().Length > 0)
                        {
                            blurImageShape.Copy();
                            var blurImageShapeCopy = thisSlide.Shapes.Paste()[1];
                            listOfBlurImageCopy.Add(blurImageShapeCopy);
                            PowerPointLabsGlobals.CopyShapePosition(blurImageShape, ref blurImageShapeCopy);
                            blurImageShapeCopy.PictureFormat.Crop.ShapeLeft = paragraph.BoundLeft - 5;
                            blurImageShapeCopy.PictureFormat.Crop.ShapeWidth = paragraph.BoundWidth + 10;
                            blurImageShapeCopy.PictureFormat.Crop.ShapeTop = paragraph.BoundTop - 5;
                            blurImageShapeCopy.PictureFormat.Crop.ShapeHeight = paragraph.BoundHeight + 10;
                            var overlayBlurShape = thisSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                            paragraph.BoundLeft - 5,
                                                            paragraph.BoundTop - 5,
                                                            paragraph.BoundWidth + 10,
                                                            paragraph.BoundHeight + 10);
                            overlayBlurShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black);
                            overlayBlurShape.Fill.Transparency = 0.85f;
                            overlayBlurShape.Line.Visible = MsoTriState.msoFalse;
                            listOfBlurImageCopy.Add(overlayBlurShape);
                            Utils.Graphics.MoveZToJustBehind(blurImageShapeCopy, shape);
                            Utils.Graphics.MoveZToJustBehind(overlayBlurShape, shape);
                            shape.Tags.Add("GotBlured", blurImageShapeCopy.Name);
                        }
                    }
                }
            }
            blurImageShape.ZOrder(MsoZOrderCmd.msoSendToBack);

            thisSlide.GetNativeSlide().Export(TextboxStyleImagePath, "JPG");

            foreach (var shape in listOfBlurImageCopy)
            {
                shape.Delete();
            }

            blurImageShape.Delete();
            overlayShape.Delete();

            //
            // Textbox style 3 starts
            //                foreach (PowerPoint.Shape shape in thisSlide.Shapes)
            //                {
            //                    if (shape.Type == MsoShapeType.msoPlaceholder
            //                        || shape.Type == MsoShapeType.msoTextBox)
            //                    {
            //                        if (shape.TextFrame.HasText == MsoTriState.msoFalse
            //                            || shape.Tags["GotHighlighted"].Trim().Length != 0)
            //                        {
            //                            continue;
            //                        }
            //                        // multiple paragraphs.. 
            //                        foreach (TextRange2 paragraph in shape.TextFrame2.TextRange.Paragraphs)
            //                        {
            //                            if (paragraph.TrimText().Length > 0)
            //                            {
            //                                var highlightShape = thisSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
            //                                                                paragraph.BoundLeft - 5,
            //                                                                paragraph.BoundTop - 5,
            //                                                                paragraph.BoundWidth + 10,
            //                                                                paragraph.BoundHeight + 10);
            //                                highlightShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black); // TODO customize
            //                                highlightShape.Line.Visible = MsoTriState.msoFalse;
            //                                Utils.Graphics.MoveZToJustBehind(highlightShape, shape);
            //                                highlightShape.Name = "PPTLabsHighlightBackgroundShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            //                                highlightShape.Tags.Add("HighlightBackground", shape.Name);
            //                                shape.Tags.Add("GotHighlighted", highlightShape.Name);
            //                            }
            //                        }
            //                    }
            //                }
            //                thisSlide.GetNativeSlide().Export(previewFile4, "JPG");
            //                PreviewList.Add(new ImageItem
            //                {
            //                    ImageFile = previewFile4
            //                });

            //
            // dont affect next time preview
            thisSlide.Delete();
        }

        private void InitImagePath()
        {
            TextboxStyleImagePath = TempPath.GetPath("textbox");
            BlurStyleImagePath = TempPath.GetPath("blur");
            DirectTextStyleImagePath = TempPath.GetPath("directtext");
        }
    }
}
