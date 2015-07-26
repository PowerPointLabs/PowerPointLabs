using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler.Preview;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.ImageSearch.Handler
{
    public class StylesHandler : PowerPointPresentation
    {
        private StyleOptions Options { get; set; }

        # region APIs

        /// <exception cref="AssumptionFailedException">
        /// throw exception when options is null
        /// </exception>
        public StylesHandler(StyleOptions options)
        {
            Assumption.Made(options != null, "options is null.");

            Path = TempPath.TempFolder;
            Name = "ImagesLabPreview";
            Options = options;
        }

        /// <exception cref="AssumptionFailedException">
        /// throw exception when ImagesLab presentation is not opened OR no selected slide.
        /// </exception>
        public PreviewInfo PreviewStyles(ImageItem source)
        {
            Assumption.Made(
                Opened && PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "ImagesLab presentation is not open OR no selected slide.");

            InitSlideSize();
            var previewInfo = new PreviewInfo();
            var handler = CreateEffectsHandler(source);

            // style: direct text
            var imageShape = ApplyDirectTextStyle(handler);
            handler.GetNativeSlide().Export(previewInfo.DirectTextStyleImagePath, "JPG");

            // style: blur
            handler.RemoveStyles(EffectsHandler.EffectName.Overlay);
            var blurImageShape = handler.ApplyBlurEffect(imageShape, Options.OverlayColor, Options.Transparency);
            handler.GetNativeSlide().Export(previewInfo.BlurStyleImagePath, "JPG");

            // style: textbox
            handler.RemoveStyles(EffectsHandler.EffectName.Overlay);
            handler.ApplyBlurTextboxEffect(blurImageShape, Options.OverlayColor, Options.Transparency);
            handler.GetNativeSlide().Export(previewInfo.TextboxStyleImagePath, "JPG");

            // style: grayscale
            handler.RemoveStyles(EffectsHandler.EffectName.Overlay);
            handler.RemoveStyles(EffectsHandler.EffectName.Blur);
            handler.ApplyGrayscaleEffect(imageShape, Options.OverlayColor, Options.Transparency);
            handler.GetNativeSlide().Export(previewInfo.GrayScaleStyleImagePath, "JPG");

            handler.Delete();
            return previewInfo;
        }

        /// <exception cref="AssumptionFailedException">
        /// throw exception when No selected slide.
        /// </exception>
        public void ApplyStyle(ImageItem source, string targetStyle)
        {
            Assumption.Made(
                PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "No selected slide.");

            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
            var effectsHandler = new EffectsHandler(currentSlide, Current, source);

            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    ApplyDirectTextStyle(effectsHandler);
                    break;
                case TextCollection.ImagesLabText.StyleNameBlur:
                    ApplyBlurStyle(effectsHandler);
                    break;
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    ApplyTextBoxStyle(effectsHandler);
                    break;
                case TextCollection.ImagesLabText.StyleNameGrayscale:
                    ApplyGrayscaleStyle(effectsHandler);
                    break;
            }
            effectsHandler.ApplyImageReference(source.ContextLink);
            ClearSelection();
        }
        # endregion

        # region Helper Funcs
        private static void ClearSelection()
        {
            var currentSelection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (currentSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                currentSelection.Unselect();
                Cursor.Current = Cursors.Default;
            }
        }

        private void ApplyGrayscaleStyle(EffectsHandler effectsHandler)
        {
            ApplyTextEffect(effectsHandler);
            effectsHandler.ApplyGrayscaleEffect(null /*no need image shape*/, Options.OverlayColor, Options.Transparency);
        }

        private void ApplyTextBoxStyle(EffectsHandler effectsHandler)
        {
            ApplyTextEffect(effectsHandler);
            effectsHandler.ApplyBackgroundEffect();
            var blurImageShape = effectsHandler.ApplyBlurEffect();
            effectsHandler.ApplyBlurTextboxEffect(blurImageShape, Options.OverlayColor, Options.Transparency);
        }

        private void ApplyBlurStyle(EffectsHandler effectsHandler)
        {
            ApplyTextEffect(effectsHandler);
            effectsHandler.ApplyBlurEffect(null /*no need image shape*/, Options.OverlayColor, Options.Transparency);
        }

        private Shape ApplyDirectTextStyle(EffectsHandler effectsHandler)
        {
            var imageShape = effectsHandler.ApplyBackgroundEffect(Options.OverlayColor, Options.Transparency);
            ApplyTextEffect(effectsHandler);
            return imageShape;
        }

        private void ApplyTextEffect(EffectsHandler effectsHandler)
        {
            if (Options.IsUseOriginalTextFormat)
            {
                effectsHandler.ApplyOriginalTextEffect();
            }
            else
            {
                effectsHandler.ApplyTextEffect(Options.GetFontFamily(), Options.FontColor, Options.FontSizeIncrease);
            }
        }

        private EffectsHandler CreateEffectsHandler(ImageItem source)
        {
            // sync layout
            var layout = PowerPointCurrentPresentationInfo.CurrentSlide.Layout;
            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);

            // sync design & theme
            newSlide.Design = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide().Design;

            return new EffectsHandler(newSlide, this, source);
        }

        private void InitSlideSize()
        {
            SlideWidth = Current.SlideWidth;
            SlideHeight = Current.SlideHeight;
        }
        #endregion
    }
}
