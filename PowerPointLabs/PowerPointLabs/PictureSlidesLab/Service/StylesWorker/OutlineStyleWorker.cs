using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class OutlineStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseOutlineStyle)
            {
                var outlineOverlayShape = designer.ApplyRectOutlineEffect(imageShape, option.FontColor, 0);
                result.Add(outlineOverlayShape);
            }
            return result;
        }
    }
}
