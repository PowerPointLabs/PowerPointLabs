using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.ImageSearch.Model;
using PowerPointLabs.ImageSearch.Slide;

namespace PowerPointLabs.ImageSearch.Presentation
{
    public class StylesPreviewPresentation : PowerPointPresentation
    {
        public const string DirectTextStyle = "directtext";
        public const string BlurStyle = "blur";
        public const string TextBoxStyle = "textbox";

        public string TextboxStyleImagePath { get; private set; }
        public string BlurStyleImagePath { get; private set; }
        public string DirectTextStyleImagePath { get; private set; }

        public StylesPreviewPresentation()
        {
            Path = TempPath.TempFolder;
            Name = "ImagesLabPreview";
        }

        public StylesPreviewSlide AddSlide(ImageItem imageItem)
        {
            if (!Opened)
            {
                return null;
            }

            var layout = PowerPointCurrentPresentationInfo.CurrentSlide.Layout;
            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);
            var previewSlide = new StylesPreviewSlide(newSlide, this, imageItem);

            Slides.Add(previewSlide);

            return previewSlide;
        }

        public void PreviewStyles(ImageItem imageItem)
        {
            InitImagePaths();
            InitSlideSize();
            
            var thisSlide = AddSlide(imageItem);

            // style: direct text
            var imageShape = thisSlide.ApplyBackgroundEffect();
            thisSlide.ApplyTextEffect();
            thisSlide.GetNativeSlide().Export(DirectTextStyleImagePath, "JPG");

            // style: blur
            var blurImageShape = thisSlide.ApplyBlurEffect(imageShape);
            thisSlide.GetNativeSlide().Export(BlurStyleImagePath, "JPG");
            thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);

            // style: textbox
            thisSlide.ApplyBlurTextboxEffect(blurImageShape);
            thisSlide.GetNativeSlide().Export(TextboxStyleImagePath, "JPG");

            thisSlide.Delete();
        }

        public void InsertStyles(ImageItem imageItem, ImageItem previewImageItem)
        {
            InitImagePaths();
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
            var thisSlide = new StylesPreviewSlide(currentSlide, Current, imageItem);
            if (previewImageItem.ImageFile.Contains(DirectTextStyle))
            {
                thisSlide.ApplyBackgroundEffect();
                thisSlide.ApplyTextEffect();
            } 
            else if (previewImageItem.ImageFile.Contains(BlurStyle))
            {
                thisSlide.ApplyTextEffect();
                thisSlide.ApplyBlurEffect();
            } 
            else if (previewImageItem.ImageFile.Contains(TextBoxStyle))
            {
                var imageShape = thisSlide.ApplyBackgroundEffect();
                thisSlide.ApplyTextEffect();
                var blurImageShape = thisSlide.ApplyBlurEffect(imageShape);
                thisSlide.RemoveStyles(StylesPreviewSlide.EffectName.Overlay);
                thisSlide.ApplyBlurTextboxEffect(blurImageShape);
            }

            var currentSelection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (currentSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                currentSelection.Unselect();
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
        }
    }
}
