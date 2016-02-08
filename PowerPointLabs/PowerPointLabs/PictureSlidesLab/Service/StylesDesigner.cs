using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Service.Preview;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker;
using PowerPointLabs.PictureSlidesLab.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    /// <summary>
    /// StylesDesigner provides APIs to preview styles
    /// and to apply the target style to a slide.
    /// 
    /// To support any new styles, create a subclass of IStyleWorker
    /// and add it to StyleWorkerFactory.
    /// </summary>
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
            
            var effectsHandler = EffectsDesigner.CreateEffectsDesignerForApply(contentSlide, 
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
            Shape imageShape;
            if (Option.IsUseSpecialEffectStyle)
            {
                imageShape = designer.ApplySpecialEffectEffect(Option.GetSpecialEffect(), isActualSize);
            }
            else // non-special-effect style, e.g. direct text style
            {
                imageShape = designer.ApplyBackgroundEffect();
            }

            var resultShapes = new List<Shape>();
            var styleWorkers = StyleWorkerFactory.GetAllStyleWorkers();
            foreach (var styleWorker in styleWorkers)
            {
                resultShapes.AddRange(
                    styleWorker.Execute(Option, designer, source, imageShape));
            }

            resultShapes.Reverse();
            SendToBack(resultShapes.ToArray());
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        # endregion

        # region Helper Funcs

        private int GetPreviewWidth()
        {
            return (int) Math.Ceiling(SlideWidth / SlideHeight * PreviewHeight);
        }

        private void SendToBack(params Shape[] shapes)
        {
            foreach (var shape in shapes)
            {
                shape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
        }

        private EffectsDesigner CreateEffectsHandlerForPreview(ImageItem source, Slide contentSlide)
        {
            // sync layout
            var layout = contentSlide.Layout;
            var newSlide = Presentation.Slides.Add(SlideCount + 1, layout);

            // sync design & theme
            newSlide.Design = contentSlide.Design;

            return EffectsDesigner.CreateEffectsDesignerForPreview(newSlide, contentSlide, SlideWidth, SlideHeight, source);
        }
        
        #endregion
    }
}
