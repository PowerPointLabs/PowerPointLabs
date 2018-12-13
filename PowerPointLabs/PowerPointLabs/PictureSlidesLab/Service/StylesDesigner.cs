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

        private EffectsDesigner EffectsDesignerForPreview { get; }

        private StyleOption Option { get; set; }

        private Settings Settings { get; set; }

        private const int PreviewHeight = 300;

        # region APIs

        public StylesDesigner(Application app = null)
        {
            Path = TempPath.TempFolder;
            Name = "PictureSlidesLabPreview" + Guid.NewGuid().ToString().Substring(0, 7) + ".pptx";
            Option = new StyleOption();
            Application = app;
            Open(withWindow: false, focus: false);
            // re use effects designer (a background slide) to
            // generate styles
            EffectsDesignerForPreview = CreateEffectsHandlerForPreview();

            AggregateCatalog catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            CompositionContainer container = new CompositionContainer(catalog);
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
            Logger.Log("PreviewApplyStyle begins");
            SetStyleOptions(option);
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;

            PreviewInfo previewInfo = new PreviewInfo();
            EffectsDesignerForPreview.PreparePreviewing(contentSlide, slideWidth, slideHeight, source);

            // use thumbnail to apply, in order to speed up
            source.BackupFullSizeImageFile = source.FullSizeImageFile;
            string backupImageFile = source.ImageFile;
            source.FullSizeImageFile = null;
            source.ImageFile = source.CroppedThumbnailImageFile ?? source.ImageFile;

            GenerateStyle(EffectsDesignerForPreview, source, isActualSize: false);

            // recover the source back
            source.FullSizeImageFile = source.BackupFullSizeImageFile;
            source.ImageFile = backupImageFile;
            EffectsDesignerForPreview.GetNativeSlide().Export(previewInfo.PreviewApplyStyleImagePath, "JPG",
                    GetPreviewWidth(), PreviewHeight);
            Logger.Log("PreviewApplyStyle done");
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
            source.BackupFullSizeImageFile = source.FullSizeImageFile;
            source.FullSizeImageFile = source.CroppedImageFile ?? source.FullSizeImageFile;
            
            EffectsDesigner effectsHandler = new EffectsDesigner(contentSlide, slideWidth, slideHeight, source);

            GenerateStyle(effectsHandler, source, isActualSize: true);

            // recover the source back
            source.FullSizeImageFile = source.BackupFullSizeImageFile;
        }

        /// <summary>
        /// process how to handle a style based on the given source and style option
        /// </summary>
        /// <param name="designer"></param>
        /// <param name="source"></param>
        /// <param name="isActualSize"></param>
        private void GenerateStyle(EffectsDesigner designer, ImageItem source, bool isActualSize)
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

            List<Shape> resultShapes = new List<Shape>();
            foreach (StylesWorker.Interface.IStyleWorker styleWorker in WorkerFactory.StyleWorkers)
            {
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
            foreach (Shape shape in shapes)
            {
                if (shape == null)
                {
                    continue;
                }

                shape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
        }

        private EffectsDesigner CreateEffectsHandlerForPreview()
        {
            Slide backgroundSlide = Presentation.Slides.Add(SlideCount + 1, PpSlideLayout.ppLayoutBlank);
            return new EffectsDesigner(backgroundSlide);
        }
        
        #endregion
    }
}
