using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class CircleStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseCircleStyle)
            {
                var circleOverlayShape = designer.ApplyCircleRingsEffect(option.CircleColor, option.CircleTransparency);
                result.Add(circleOverlayShape);
            }
            return result;
        }
    }
}
