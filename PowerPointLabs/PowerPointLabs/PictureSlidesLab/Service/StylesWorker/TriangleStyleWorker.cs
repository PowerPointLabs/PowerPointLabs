using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class TriangleStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseTriangleStyle)
            {
                var triangleOverlayShape = designer.ApplyTriangleEffect(option.TriangleColor, option.FontColor,
                    option.TriangleTransparency);
                result.Add(triangleOverlayShape);
            }
            return result;
        }
    }
}
