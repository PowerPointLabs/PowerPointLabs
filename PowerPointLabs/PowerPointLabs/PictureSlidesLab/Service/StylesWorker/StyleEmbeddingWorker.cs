using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 1)]
    class StyleEmbeddingWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            // in previewing 
            if (source.FullSizeImageFile == null)
            {
                return new List<Shape>();
            }
            // store style options information into original image shape
            // return original image and cropped image
            return designer.EmbedStyleOptionsInformation(
                source.BackupFullSizeImageFile, source.FullSizeImageFile,
                source.ContextLink, source.Source, source.Rect, option);
        }
    }
}
