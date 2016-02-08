using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Service.Preview;
using PowerPointLabs.PictureSlidesLab.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    public sealed class StylesDesigner : PowerPointPresentation, IStylesDesigner
    {
        private StyleOption Option { get; set; }

        private const int PreviewHeight = 300;

        # region APIs

        public StylesDesigner(Application app = null)
        {
            Path = TempPath.TempFolder;
            Name = "PictureSlidesLabPreview";
            Option = new StyleOption();
            Application = app;
            Open(withWindow: false, focus: false);
        }

        public void SetStyleOptions(StyleOption opt)
        {
            Option = opt;
        }

        public void CleanUp()
        {
            Close();
        }

        public PreviewInfo PreviewApplyStyle(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight, StyleOption option)
        {
            SetStyleOptions(option);
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;

            var previewInfo = new PreviewInfo();
            var handler = CreateEffectsHandlerForPreview(source, contentSlide);

            // use thumbnail to apply, in order to speed up
            var fullSizeImgPath = source.FullSizeImageFile;
            var originalThumbnail = source.ImageFile;
            source.FullSizeImageFile = null;
            source.ImageFile = source.CroppedThumbnailImageFile ?? source.ImageFile;

            ApplyStyle(handler, source, isActualSize: false);

            // recover the source back
            source.FullSizeImageFile = fullSizeImgPath;
            source.ImageFile = originalThumbnail;
            handler.GetNativeSlide().Export(previewInfo.PreviewApplyStyleImagePath, "JPG",
                    GetPreviewWidth(), PreviewHeight);

            handler.Delete();
            return previewInfo;
        }
        
        public void ApplyStyle(ImageItem source, Slide contentSlide,
            float slideWidth, float slideHeight, StyleOption option = null)
        {
            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
            }
            if (option != null)
            {
                SetStyleOptions(option);
            }

            // try to use cropped/adjusted image to apply
            var fullsizeImage = source.FullSizeImageFile;
            source.FullSizeImageFile = source.CroppedImageFile ?? source.FullSizeImageFile;
            source.OriginalImageFile = fullsizeImage;
            
            var effectsHandler = new EffectsDesigner(contentSlide, 
                slideWidth, slideHeight, source);

            ApplyStyle(effectsHandler, source, isActualSize: true);

            // recover the source back
            source.FullSizeImageFile = fullsizeImage;
            source.OriginalImageFile = null;
        }

        /// <summary>
        /// process how to handle a style based on the given source and style option
        /// </summary>
        /// <param name="designer"></param>
        /// <param name="source"></param>
        /// <param name="isActualSize"></param>
        private void ApplyStyle(EffectsDesigner designer, ImageItem source, bool isActualSize)
        {
            // TODO refactor this method
            designer.ApplyPseudoTextWhenNoTextShapes();

            if (Option.IsUseBannerStyle 
                && (Option.TextBoxPosition == 4/*left*/
                    || Option.TextBoxPosition == 5/*centered*/
                    || Option.TextBoxPosition == 6/*right*/))
            {
                designer.ApplyTextWrapping();
            }
            else if (Option.IsUseCircleStyle
                     || Option.IsUseOutlineStyle)
            {
                designer.ApplyTextWrapping();
            }
            else
            {
                designer.RecoverTextWrapping();
            }

            ApplyTextEffect(designer);
            designer.ApplyTextGlowEffect(Option.IsUseTextGlow, Option.TextGlowColor);

            // store style options information into original image shape
            // return original image and cropped image
            var metaImages = designer.EmbedStyleOptionsInformation(
                source.OriginalImageFile, source.FullSizeImageFile, 
                source.ContextLink, source.Rect, Option);
            Shape originalImage = null;
            Shape croppedImage = null;
            if (metaImages.Count == 2)
            {
                originalImage = metaImages[0];
                croppedImage = metaImages[1];
            }

            Shape imageShape;
            if (Option.IsUseSpecialEffectStyle)
            {
                imageShape = designer.ApplySpecialEffectEffect(Option.GetSpecialEffect(), isActualSize);
            }
            else // Direct Text style
            {
                imageShape = designer.ApplyBackgroundEffect();
            }

            Shape backgroundOverlayShape = null;
            if (Option.IsUseOverlayStyle)
            {
                backgroundOverlayShape = designer.ApplyOverlayEffect(Option.OverlayColor, Option.Transparency);
            }

            Shape blurImageShape = null;
            if (Option.IsUseBlurStyle)
            {
                blurImageShape = Option.IsUseSpecialEffectStyle
                    ? designer.ApplyBlurEffect(source.SpecialEffectImageFile, Option.BlurDegree)
                    : designer.ApplyBlurEffect(degree: Option.BlurDegree);
            }

            Shape bannerOverlayShape = null;
            if (Option.IsUseBannerStyle)
            {
                bannerOverlayShape = ApplyBannerStyle(designer, imageShape);
            }

            if (Option.IsUseTextBoxStyle)
            {
                designer.ApplyTextboxEffect(Option.TextBoxColor, Option.TextBoxTransparency);
            }

            Shape outlineOverlayShape = null;
            if (Option.IsUseOutlineStyle)
            {
                outlineOverlayShape = designer.ApplyRectOutlineEffect(imageShape, Option.FontColor, 0);
            }

            Shape frameOverlayShape = null;
            if (Option.IsUseFrameStyle)
            {
                frameOverlayShape = designer.ApplyAlbumFrameEffect(Option.FrameColor, Option.FrameTransparency);
            }

            Shape circleOverlayShape = null;
            if (Option.IsUseCircleStyle)
            {
                circleOverlayShape = designer.ApplyCircleRingsEffect(Option.CircleColor, Option.CircleTransparency);
            }

            Shape triangleOverlayShape = null;
            if (Option.IsUseTriangleStyle)
            {
                triangleOverlayShape = designer.ApplyTriangleEffect(Option.TriangleColor, Option.FontColor,
                    Option.TriangleTransparency);
            }

            SendToBack(
                triangleOverlayShape,
                circleOverlayShape,
                frameOverlayShape,
                outlineOverlayShape,
                bannerOverlayShape,
                backgroundOverlayShape,
                blurImageShape,
                imageShape,
                croppedImage,
                originalImage);

            designer.ApplyImageReference(source.ContextLink);
            if (Option.IsInsertReference)
            {
                designer.ApplyImageReferenceInsertion(source.ContextLink, Option.GetFontFamily(), Option.FontColor,
                    Option.CitationFontSize, Option.ImageReferenceTextBoxColor, Option.GetCitationTextBoxAlignment());
            }
        }

        # endregion

        # region Helper Funcs

        private int GetPreviewWidth()
        {
            return (int)(SlideWidth / SlideHeight * PreviewHeight);
        }

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

        private Shape ApplyBannerStyle(EffectsDesigner effectsDesigner, Shape imageShape)
        {
            return effectsDesigner.ApplyRectBannerEffect(Option.GetBannerDirection(), Option.GetTextBoxPosition(),
                        imageShape, Option.BannerColor, Option.BannerTransparency);
        }

        private void ApplyTextEffect(EffectsDesigner effectsDesigner)
        {
            if (Option.IsUseTextFormat)
            {
                effectsDesigner.ApplyTextEffect(Option.GetFontFamily(), Option.FontColor, Option.FontSizeIncrease);
                effectsDesigner.ApplyTextPositionAndAlignment(Option.GetTextBoxPosition(), Option.GetTextAlignment());
            }
            else
            {
                effectsDesigner.ApplyOriginalTextEffect();
                effectsDesigner.ApplyTextPositionAndAlignment(Position.Original, Alignment.Auto);
            }
            
        }

        private EffectsDesigner CreateEffectsHandlerForPreview(ImageItem source, Slide contentSlide)
        {
            // sync layout
            var layout = contentSlide.Layout;
            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);

            // sync design & theme
            newSlide.Design = contentSlide.Design;

            return new EffectsDesigner(newSlide, contentSlide, SlideWidth, SlideHeight, source);
        }
        
        #endregion
    }
}
