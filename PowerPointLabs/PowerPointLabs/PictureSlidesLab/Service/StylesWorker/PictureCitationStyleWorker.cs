using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 12)]
    class PictureCitationStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            designer.ApplyImageReference(source.Source);
            if (settings != null && settings.IsInsertCitation)
            {
                designer.ApplyImageReferenceInsertion(source.Source, "Calibri", settings.CitationFontColor,
                    settings.CitationFontSize, 
                    settings.IsUseCitationTextBox ? settings.CitationTextBoxColor : "", 
                    settings.GetCitationTextBoxAlignment());
            }
            else if (option.IsInsertReference)
            {
                designer.ApplyImageReferenceInsertion(source.Source, option.GetFontFamily(), option.FontColor,
                    option.CitationFontSize, option.ImageReferenceTextBoxColor, option.GetCitationTextBoxAlignment());
            }
            return new List<Shape>();
        }
    }
}
