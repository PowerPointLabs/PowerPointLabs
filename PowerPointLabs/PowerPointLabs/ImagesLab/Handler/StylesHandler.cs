using Microsoft.Office.Core;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Handler.Effect;
using PowerPointLabs.ImagesLab.Handler.Preview;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ImagesLab.Handler
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
        /// throw exception when ImagesLab presentation is not open OR no selected slide.
        /// </exception>
        public PreviewInfo PreviewApplyStyle(ImageItem source)
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
            source.FullSizeImageFile = null;
            source.ImageFile = source.CroppedThumbnailImageFile ?? source.ImageFile;

            ApplyStyle(handler, source, isActualSize: false);

            // recover it back
            source.FullSizeImageFile = fullSizeImgPath;
            source.ImageFile = originalThumbnail;
            handler.GetNativeSlide().Export(previewInfo.PreviewApplyStyleImagePath, "JPG",
                    GetPreviewWidth(), PreviewHeight);

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

        /// <summary>
        /// process how to handle a style based on the given source and style option
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="source"></param>
        /// <param name="isActualSize"></param>
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

            // store style options information into original image shape
            // return original image and cropped image
            var metaImages = handler.EmbedStyleOptionsInformation(
                source.OriginalImageFile, source.FullSizeImageFile, source.Rect, Options);
            Shape originalImage = null;
            Shape croppedImage = null;
            if (metaImages.Count == 2)
            {
                originalImage = metaImages[0];
                croppedImage = metaImages[1];
            }

            Shape imageShape;
            if (Options.IsUseSpecialEffectStyle)
            {
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
                blurImageShape = Options.IsUseSpecialEffectStyle
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
                handler.ApplyTextboxEffect(Options.TextBoxColor, Options.TextBoxTransparency);
            }

            Shape outlineOverlayShape = null;
            if (Options.IsUseOutlineStyle)
            {
                outlineOverlayShape = handler.ApplyRectOutlineEffect(imageShape, Options.FontColor, 0);
            }

            Shape frameOverlayShape = null;
            if (Options.IsUseFrameStyle)
            {
                frameOverlayShape = handler.ApplyAlbumFrameEffect(Options.FrameColor, Options.FrameTransparency);
            }

            SendToBack(
                    frameOverlayShape,
                    outlineOverlayShape,
                    bannerOverlayShape,
                    backgroundOverlayShape,
                    blurImageShape,
                    imageShape,
                    croppedImage,
                    originalImage);

            handler.ApplyImageReference(source.ContextLink);
            if (Options.IsInsertReference)
            {
                handler.ApplyImageReferenceInsertion(source.ContextLink, Options.GetFontFamily(), Options.FontColor);
            }
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
            return effectsHandler.ApplyRectBannerEffect(Options.GetBannerDirection(), Options.GetTextBoxPosition(),
                        imageShape, Options.BannerColor, Options.BannerTransparency);
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
