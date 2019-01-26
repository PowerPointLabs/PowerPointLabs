using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 5)]
    class TextBoxStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            if (option.IsUseTextBoxStyle)
            {
                designer.ApplyTextboxEffect(option.TextBoxColor, option.TextBoxTransparency, option.FontSizeIncrease);
            }
            return new List<Shape>();
        }
    }
}
