using ImageProcessor.Imaging.Filters;
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

        public StylesHandler()
        {
            Path = TempPath.TempFolder;
            Name = "ImagesLabPreview";
            Options = new StyleOptions();
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

            // use (cropped) thumbnail to apply, in order to speed up
            var fullSizeImgPath = source.FullSizeImageFile;
            var originalThumbnail = source.ImageFile;
            source.FullSizeImageFile = null;
            source.ImageFile = source.CroppedThumbnailImageFile ?? source.ImageFile;

            PreviewStyles(handler, previewInfo);

            source.ImageFile = originalThumbnail;
            source.FullSizeImageFile = fullSizeImgPath;

            handler.Delete();
            return previewInfo;
        }

        /// <exception cref="AssumptionFailedException">
        /// throw exception when ImagesLab presentation is not open OR no selected slide.
        /// </exception>
        public PreviewInfo PreviewApplyStyle(ImageItem source, bool isActualSize = false)
        {
            Assumption.Made(
                Opened && PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "ImagesLab presentation is not open OR no selected slide.");

            InitSlideSize();
            var previewInfo = new PreviewInfo();
            var handler = CreateEffectsHandler(source);

            // use thumbnail to apply, in order to speed up
            var fullSizeImgPath = source.FullSizeImageFile;
            var originalThumbnail = source.ImageFile;
            if (!isActualSize)
            {
                source.FullSizeImageFile = null;
                source.ImageFile = source.CroppedThumbnailImageFile ?? source.ImageFile;
            }
            else
            {
                source.FullSizeImageFile = source.CroppedImageFile ?? source.FullSizeImageFile;
            }

            ApplyStyle(handler, source, isActualSize);

            source.FullSizeImageFile = fullSizeImgPath;
            source.ImageFile = originalThumbnail;

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
        public void ApplyStyle(ImageItem source)
        {
            Assumption.Made(
                PowerPointCurrentPresentationInfo.CurrentSlide != null,
                "No selected slide.");

            Globals.ThisAddIn.Application.StartNewUndoEntry();

            // try to use cropped/adjusted image to apply
            var fullsizeImage = source.FullSizeImageFile;
            source.FullSizeImageFile = source.CroppedImageFile ?? source.FullSizeImageFile;
            source.OriginalImageFile = fullsizeImage;

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
            var effectsHandler = new EffectsHandler(currentSlide, Current, source);

            ApplyStyle(effectsHandler, source, isActualSize:true);

            source.FullSizeImageFile = fullsizeImage;
            source.OriginalImageFile = null;
        }

        private int GetPreviewWidth()
        {
            return (int)(SlideWidth / SlideHeight * PreviewHeight);
        }

        private void ApplyStyle(EffectsHandler handler, ImageItem source, bool isActualSize)
        {
            if (Options.IsUseBannerStyle 
                && (Options.TextBoxPosition == 4/*left*/
                    || Options.TextBoxPosition == 5/*centered*/
                    || Options.TextBoxPosition == 6/*right*/))
            {
                handler.ApplyTextWrapping();
            }
            ApplyTextEffect(handler);

            var isSpecialEffectStyle = false;

            // store style options information into original image shape
            // return original image and cropped image
            var metaImages = handler.EmbedStyleOptionsInformation(
                source.OriginalImageFile, source.FullSizeImageFile, source.Rect, Options);

            Shape imageShape;
            if (Options.IsUseSpecialEffectStyle)
            {
                isSpecialEffectStyle = true;
                imageShape = handler.ApplySpecialEffectEffect(Options.GetSpecialEffect(), isActualSize, Options.ImageOffset);
            }
            else // Direct Text style
            {
                imageShape = handler.ApplyBackgroundEffect(Options.ImageOffset);
            }

            Shape backgroundOverlayShape = null;
            if (Options.IsUseOverlayStyle)
            {
                backgroundOverlayShape = handler.ApplyOverlayEffect(Options.OverlayColor, Options.Transparency);
            }

            Shape blurImageShape = null;
            if (Options.IsUseBlurStyle)
            {
                blurImageShape = isSpecialEffectStyle
                    ? handler.ApplyBlurEffect(source.SpecialEffectImageFile, Options.BlurDegree, Options.ImageOffset)
                    : handler.ApplyBlurEffect(degree: Options.BlurDegree, offset: Options.ImageOffset);
            }

            Shape bannerOverlayShape = null;
            if (Options.IsUseBannerStyle)
            {
                bannerOverlayShape = ApplyBannerStyle(handler, imageShape);
            }

            if (Options.IsUseTextBoxStyle)
            {
                handler.ApplyTextboxEffect(Options.TextBoxOverlayColor, Options.TextBoxTransparency);
            }

            if (metaImages.Count == 2)
            {
                SendToBack(bannerOverlayShape, backgroundOverlayShape, blurImageShape, imageShape, metaImages[1],
                    metaImages[0]);
            }
            else
            {
                SendToBack(bannerOverlayShape, backgroundOverlayShape, blurImageShape, imageShape);
            }

            handler.ApplyImageReference(source.ContextLink);
            if (Options.IsInsertReference)
            {
                handler.ApplyImageReferenceInsertion(source.ContextLink, Options.GetFontFamily(), Options.FontColor);
            }
        }

        // generate preview-style images, without using any style options
        private void PreviewStyles(EffectsHandler handler, PreviewInfo previewInfo)
        {
            // style: direct text
            var imageShape = handler.ApplyBackgroundEffect(offset: 0);
            handler.ApplyTextEffect("Calibri", "#FFFFFF", 0);
            handler.ApplyTextPositionAndAlignment(Position.Centre, Alignment.Auto);
            handler.GetNativeSlide().Export(previewInfo.DirectTextStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: blur
            var blurImageShape = handler.ApplyBlurEffect(offset: 0);
            SendToBack(blurImageShape, imageShape);
            handler.GetNativeSlide().Export(previewInfo.BlurStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: textbox
            handler.RemoveEffect(EffectName.Blur);
            handler.ApplyTextPositionAndAlignment(Position.BottomLeft, Alignment.Left);
            handler.ApplyTextEffect("Calibri", "#FFD700", 0);
            handler.ApplyTextboxEffect("#000000", 25);
            handler.GetNativeSlide().Export(previewInfo.TextboxStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: banner
            handler.RemoveEffect(EffectName.TextBox);
            handler.ApplyTextEffect("Calibri", "#FFD700", 0);
            handler.ApplyRectBannerEffect(BannerDirection.Horizontal, Position.BottomLeft,
                        imageShape, "#000000", 25);
            handler.GetNativeSlide().Export(previewInfo.BannerStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: overlay
            handler.RemoveEffect(EffectName.Banner);
            handler.ApplyTextEffect("Calibri", "#FFFFFF", 0);
            handler.ApplyTextPositionAndAlignment(Position.Centre, Alignment.Left);
            handler.ApplySpecialEffectEffect(MatrixFilters.GreyScale, imageShape, "#007FFF"/*Blue*/, transparency: 35, offset: 0);
            handler.GetNativeSlide().Export(previewInfo.OverlayStyleImagePath, "JPG",
                GetPreviewWidth(), PreviewHeight);

            // style: special effect
            handler.RemoveEffect(EffectName.Overlay);
            handler.RemoveEffect(EffectName.SpecialEffect);
            handler.ApplySpecialEffectEffect(MatrixFilters.GreyScale, imageShape, "#000000", transparency: 100, offset: 0);
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

        private Shape ApplyBannerStyle(EffectsHandler effectsHandler, Shape imageShape)
        {
            switch (Options.GetBannerShape())
            {
                case BannerShape.Rectangle:
                    return effectsHandler.ApplyRectBannerEffect(Options.GetBannerDirection(), Options.GetTextBoxPosition(),
                        imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
                case BannerShape.Circle:
                    return effectsHandler.ApplyCircleBannerEffect(imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
                case BannerShape.RectangleOutline:
                    return effectsHandler.ApplyRectOutlineEffect(imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
                default:
                    return effectsHandler.ApplyCircleOutlineEffect(imageShape, Options.BannerOverlayColor, Options.BannerTransparency);
            }
        }

        private void ApplyTextEffect(EffectsHandler effectsHandler)
        {
            if (Options.IsUseTextFormat)
            {
                effectsHandler.ApplyTextEffect(Options.GetFontFamily(), Options.FontColor, Options.FontSizeIncrease);
                effectsHandler.ApplyTextPositionAndAlignment(Options.GetTextBoxPosition(), Options.GetTextBoxAlignment());
            }
            else
            {
                effectsHandler.ApplyOriginalTextEffect();
                effectsHandler.ApplyTextPositionAndAlignment(Position.Original, Alignment.Auto);
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
