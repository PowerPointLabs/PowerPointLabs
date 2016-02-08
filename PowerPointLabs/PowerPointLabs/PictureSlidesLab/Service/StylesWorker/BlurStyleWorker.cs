using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class BlurStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseBlurStyle)
            {
                var blurImageShape = option.IsUseSpecialEffectStyle
                    ? designer.ApplyBlurEffect(source.SpecialEffectImageFile, option.BlurDegree)
                    : designer.ApplyBlurEffect(degree: option.BlurDegree);
                result.Add(blurImageShape);
            }
            return result;
        }
    }
}
