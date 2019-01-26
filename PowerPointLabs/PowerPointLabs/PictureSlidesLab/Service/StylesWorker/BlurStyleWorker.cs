using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 2)]
    class BlurStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            List<Shape> result = new List<Shape>();
            if (option.IsUseBlurStyle)
            {
                Shape blurImageShape = option.IsUseSpecialEffectStyle
                    ? designer.ApplyBlurEffect(source.SpecialEffectImageFile, option.BlurDegree)
                    : designer.ApplyBlurEffect(degree: option.BlurDegree);
                result.Add(blurImageShape);
            }
            return result;
        }
    }
}
