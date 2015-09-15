using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler.Effect;
using PowerPointLabs.ImageSearch.Handler.Preview;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ImageSearch.Handler
{
    public class StylesHandler : PowerPointPresentation
    {
        private StyleOptions Options { get; set; }

        private const int PreviewHeight = 300;

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

        public void SetStyleOptions(StyleOptions opt)
        {
            Options = opt;
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

            PreviewStyles(source, handler, previewInfo);

            handler.Delete();
            return previewInfo;
        }

        /// <exception cref="AssumptionFailedException">
        /// throw exception when ImagesLab presentation is not open OR no selected slide.
        /// </exception>
        public PreviewInfo PreviewApplyStyle(ImageItem source, IList<string> targetStyles, bool isActualSize = false)
        {
            Assumption.Made(
                Opened && PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "ImagesLab presentation is not open OR no selected slide.");

            InitSlideSize();
            var previewInfo = new PreviewInfo();
            var handler = CreateEffectsHandler(source);

            ApplyStyle(handler, source, targetStyles, isActualSize);

            if (isActualSize)
            {
                handler.GetNativeSlide().Export(previewInfo.PreviewApplyStyleImagePath, "JPG");
            }
            else
            {
                handler.GetNativeSlide().Export(previewInfo.PreviewApplyStyleImagePath, "JPG",
                    GetPreviewWidth(), PreviewHeight);
            }

            handler.Delete();
            return previewInfo;
        }

        /// <exception cref="AssumptionFailedException">
        /// throw exception when No selected slide.
        /// </exception>
        public void ApplyStyle(ImageItem source, IList<string> targetStyles)
        {
            Assumption.Made(
                PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "No selected slide.");

            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
            var effectsHandler = new EffectsHandler(currentSlide, Current, source);

            ApplyStyle(effectsHandler, source, targetStyles, isActualSize:true);
        }

        private int GetPreviewWidth()
        {
            return (int)(SlideWidth / SlideHeight * PreviewHeight);
        }

        private void ApplyStyle(EffectsHandler handler, ImageItem source, IList<string> targetStyles, bool isActualSize)
        {
            ApplyTextEffect(handler);

            var isSpecialEffectStyle = false;

            Shape imageShape;
            if (HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameSpecialEffect))
            {
                isSpecialEffectStyle = true;
                imageShape = handler.ApplySpecialEffectEffect(Options.GetSpecialEffect(), isActualSize);
            }
            else // Direct Text style
            {
                imageShape = handler.ApplyBackgroundEffect();
            }
            Shape backgroundOverlayShape = handler.ApplyOverlayEffect(Options.OverlayColor, Options.Transparency);

            Shape blurImageShape = null;
            if (HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameBlur))
            {
                blurImageShape = isSpecialEffectStyle
                    ? handler.ApplyBlurEffect(source.SpecialEffectImageFile)
                    : handler.ApplyBlurEffect();
            }

            Shape bannerOverlayShape = null;
            if (HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameBanner))
            {
                bannerOverlayShape = ApplyBannerStyle(handler, imageShape);
            }

            Shape outlineOverlayShape = null;
            if (HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameOutline))
            {
                outlineOverlayShape = ApplyOutlineStyle(handler, imageShape);
            }

            if (HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameTextBox))
            {
                handler.ApplyTextboxEffect(Options.TextBoxOverlayColor, Options.TextBoxTransparency);
            }

            SendToBack(outlineOverlayShape, bannerOverlayShape, backgroundOverlayShape, blurImageShape, imageShape);

            handler.ApplyImageReference(source.ContextLink);
            if (Options.IsInsertReference)
            {
                handler.ApplyImageReferenceInsertion(source.ContextLink, Options.GetFontFamily(), Options.FontColor);
            }
        }

        private void PreviewStyles(ImageItem source, EffectsHandler handler, PreviewInfo previewInfo)
        {
            // style: direct text
            var imageShape = handler.ApplyBackgroundEffect();
            var overlayShape = handler.ApplyOverlayEffect(Options.OverlayColor, Options.Transparency);
            SendToBack(overlayShape, imageShape);
            ApplyTextEffect(handler);
            if (Options.IsInsertReference)
            {
                handler.ApplyImageReferenceInsertion(source.ContextLink, Options.GetFontFamily(), Options.FontColor);
            }
            handler.GetNativeSlide().Export(previewInfo.DirectTextStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: blur
            var blurImageShape = handler.ApplyBlurEffect();
            SendToBack(overlayShape, blurImageShape, imageShape);
            handler.GetNativeSlide().Export(previewInfo.BlurStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: textbox
            handler.RemoveEffect(EffectName.Blur);
            handler.ApplyTextboxEffect(Options.TextBoxOverlayColor, Options.TextBoxTransparency);
            handler.GetNativeSlide().Export(previewInfo.TextboxStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: banner
            handler.RemoveEffect(EffectName.TextBox);
            ApplyBannerStyle(handler, imageShape);
            handler.GetNativeSlide().Export(previewInfo.BannerStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: outline
            handler.RemoveEffect(EffectName.Banner);
            ApplyOutlineStyle(handler, imageShape);
            handler.GetNativeSlide().Export(previewInfo.OutlineStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: special effect
            handler.RemoveEffect(EffectName.Banner);
            handler.ApplySpecialEffectEffect(Options.GetSpecialEffect(), imageShape, Options.OverlayColor, Options.Transparency);
            handler.GetNativeSlide().Export(previewInfo.SpecialEffectStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);
        }

        # endregion

        # region Helper Funcs

        private void SendToBack(params Shape[] shapes)
        {
            foreach (var shape in shapes)
            {
                SendToBack(shape);
            }
        }

        private void SendToBack(Shape shape)
        {
            if (shape != null)
            {
                shape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
        }

        private bool HasStyle(IList<string> targetStyles, string style)
        {
            return targetStyles.Any(targetStyle => targetStyle == style);
        }

        private Shape ApplyBannerStyle(EffectsHandler effectsHandler, Shape imageShape)
        {
            switch (Options.GetBannerShape())
            {
                case BannerShape.Rectangle:
                    return effectsHandler.ApplyRectBannerEffect(Options.GetBannerDirection(), Options.GetTextBoxPosition(),
                        imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
                default:
                    return effectsHandler.ApplyCircleBannerEffect(imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
            }
        }

        private Shape ApplyOutlineStyle(EffectsHandler effectsHandler, Shape imageShape)
        {
            switch (Options.GetOutlineShape())
            {
                case OutlineShape.RectangleOutline:
                    return effectsHandler.ApplyRectOutlineEffect(imageShape, Options.OutlineOverlayColor,
                        Options.OutlineTransparency);
                default:
                    return effectsHandler.ApplyCircleOutlineEffect(imageShape, Options.OutlineOverlayColor,
                        Options.OutlineTransparency);
            }
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
            effectsHandler.ApplyTextPositionAndAlignment(Options.GetTextBoxPosition(), Options.GetTextBoxAlignment());
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
