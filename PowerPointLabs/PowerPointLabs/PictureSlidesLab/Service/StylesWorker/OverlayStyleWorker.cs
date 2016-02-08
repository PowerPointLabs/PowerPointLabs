using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class OverlayStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseOverlayStyle)
            {
                var backgroundOverlayShape = designer.ApplyOverlayEffect(option.OverlayColor, option.Transparency);
                result.Add(backgroundOverlayShape);
            }
            return result;
        }
    }
}
