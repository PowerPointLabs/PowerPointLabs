using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class StyleEmbeddingWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source,
            Shape imageShape)
        {
            // store style options information into original image shape
            // return original image and cropped image
            return designer.EmbedStyleOptionsInformation(
                source.OriginalImageFile, source.FullSizeImageFile,
                source.ContextLink, source.Rect, option);
        }
    }
}
