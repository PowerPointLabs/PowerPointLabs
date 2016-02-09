using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 10)]
    class PictureCitationStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            designer.ApplyImageReference(source.ContextLink);
            if (option.IsInsertReference)
            {
                designer.ApplyImageReferenceInsertion(source.ContextLink, option.GetFontFamily(), option.FontColor,
                    option.CitationFontSize, option.ImageReferenceTextBoxColor, option.GetCitationTextBoxAlignment());
            }
            return new List<Shape>();
        }
    }
}
