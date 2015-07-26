using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Slide;

namespace PowerPointLabs.ImageSearch.Presentation
{
    public class StylesPreviewPresentation : PowerPointPresentation
    {
        public const string DirectTextStyle = "directtext";
        public const string BlurStyle = "blur";
        public const string TextBoxStyle = "textbox";
        public const string GrayscaleStyle = "grayscale";

        public string GrayScaleStyleImagePath { get; private set; }
        public string TextboxStyleImagePath { get; private set; }
        public string BlurStyleImagePath { get; private set; }
        public string DirectTextStyleImagePath { get; private set; }
        private StyleOptions Options { get; set; }

        public StylesPreviewPresentation(StyleOptions options)
        {
            Path = TempPath.TempFolder;
            Name = "ImagesLabPreview";
            Options = options;
        }

        public StylesPreviewSlide AddSlide(ImageItem imageItem)
        {
            // Assumption: current slide is not null
            if (!Opened)
            {
                return null;
            }

            // sync layout
            var layout = PowerPointCurrentPresentationInfo.CurrentSlide.Layout;
            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);

            // sync design & theme
            newSlide.Design = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide().Design;

            return new StylesPreviewSlide(newSlide, this, imageItem);
        }

        public void PreviewStyles(ImageItem imageItem)
        {
            InitImagePaths();
            InitSlideSize();
            
            var thisSlide = AddSlide(imageItem);

            // style: direct text
            var imageShape = thisSlide.ApplyBackgroundEffect(Options.OverlayColor, Options.Transparency);
            ApplyTextEffect(thisSlide);
            thisSlide.GetNativeSlide().Export(DirectTextStyleImagePath, "JPG");
            thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);

            // style: blur
            var blurImageShape = thisSlide.ApplyBlurEffect(imageShape, Options.OverlayColor, Options.Transparency);
            thisSlide.GetNativeSlide().Export(BlurStyleImagePath, "JPG");
            thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);

            // style: textbox
            thisSlide.ApplyBlurTextboxEffect(blurImageShape, Options.OverlayColor, Options.Transparency);
            thisSlide.GetNativeSlide().Export(TextboxStyleImagePath, "JPG");
            thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);
            thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Blur);

            // style: grayscale
            thisSlide.ApplyGrayscaleEffect(imageShape, Options.OverlayColor, Options.Transparency);
            thisSlide.GetNativeSlide().Export(GrayScaleStyleImagePath, "JPG");

            thisSlide.Delete();
        }

        public void InsertStyles(ImageItem imageItem, ImageItem previewImageItem)
        {
            // TODO refactor options thing
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            InitImagePaths();
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
            var thisSlide = new StylesPreviewSlide(currentSlide, Current, imageItem);
            if (previewImageItem.ImageFile.Contains(DirectTextStyle))
            {
                thisSlide.ApplyBackgroundEffect(Options.OverlayColor, Options.Transparency);
                ApplyTextEffect(thisSlide);
            }
            else if (previewImageItem.ImageFile.Contains(BlurStyle))
            {
                ApplyTextEffect(thisSlide);
                thisSlide.ApplyBlurEffect(null/*no need clear image shape*/, Options.OverlayColor, Options.Transparency);
            } 
            else if (previewImageItem.ImageFile.Contains(TextBoxStyle))
            {
                var imageShape = thisSlide.ApplyBackgroundEffect(Options.OverlayColor, Options.Transparency);
                ApplyTextEffect(thisSlide);
                var blurImageShape = thisSlide.ApplyBlurEffect(imageShape, Options.OverlayColor, Options.Transparency);
                thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);
                thisSlide.ApplyBlurTextboxEffect(blurImageShape, Options.OverlayColor, Options.Transparency);
            }
            else if (previewImageItem.ImageFile.Contains(GrayscaleStyle))
            {
                ApplyTextEffect(thisSlide);
                thisSlide.ApplyGrayscaleEffect(null/*no need clear image shape*/, Options.OverlayColor, Options.Transparency);
            }

            thisSlide.ApplyImageReference(imageItem.ContextLink);

            var currentSelection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (currentSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                currentSelection.Unselect();
                Cursor.Current = Cursors.Default;
            }
        }

        private void ApplyTextEffect(StylesPreviewSlide thisSlide)
        {
            if (Options.IsUseOriginalTextFormat)
            {
                thisSlide.ApplyOriginalTextEffect();
            }
            else
            {
                thisSlide.ApplyTextEffect(Options.GetFontFamily(), Options.FontColor, Options.FontSizeIncrease);
            }
        }

        private void InitSlideSize()
        {
            SlideWidth = Current.SlideWidth;
            SlideHeight = Current.SlideHeight;
        }

        private void InitImagePaths()
        {
            TextboxStyleImagePath = TempPath.GetPath(TextBoxStyle);
            BlurStyleImagePath = TempPath.GetPath(BlurStyle);
            DirectTextStyleImagePath = TempPath.GetPath(DirectTextStyle);
            GrayScaleStyleImagePath = TempPath.GetPath(GrayscaleStyle);
        }
    }
}
