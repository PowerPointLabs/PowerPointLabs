using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Service.Preview;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Factory;
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
        [Import(typeof(StyleWorkerFactory))]
        private StyleWorkerFactory WorkerFactory { get; set; }

        private StyleOption Option { get; set; }

        private Settings Settings { get; set; }

        private const int PreviewHeight = 300;

        # region APIs

        public StylesDesigner(Application app = null)
        {
            Path = TempPath.TempFolder;
            Name = "PictureSlidesLabPreview" + Guid.NewGuid().ToString().Substring(0, 7);
            Option = new StyleOption();
            Application = app;
            OpenInBackground();

            var catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            var container = new CompositionContainer(catalog);
            container.ComposeParts(this);
        }

        public void SetStyleOptions(StyleOption opt)
        {
            Option = opt;
        }

        public void SetSettings(Settings settings)
        {
            Settings = settings;
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
            Logger.Log("Generate style " + Option.StyleName);
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
            foreach (var styleWorker in WorkerFactory.StyleWorkers)
            {
                Logger.Log("Executing worker " + styleWorker.GetType().Name);
                resultShapes.AddRange(
                    styleWorker.Execute(Option, designer, source, imageShape, Settings));
            }
            // Those workers executed at the beginning will have the output
            // put at the back.
            resultShapes.Reverse();
            SendToBack(resultShapes.ToArray());
            imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            Logger.Log("Complete generating style " + Option.StyleName);
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
                if (shape == null)
                {
                    continue;
                }

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
