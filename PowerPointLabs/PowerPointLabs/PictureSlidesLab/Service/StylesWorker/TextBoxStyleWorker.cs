using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class TextBoxStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            if (option.IsUseTextBoxStyle)
            {
                designer.ApplyTextboxEffect(option.TextBoxColor, option.TextBoxTransparency);
            }
            return new List<Shape>();
        }
    }
}
