using System.Collections.Generic;
using System.ComponentModel.Composition;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 11)]
    class FrostedGlassTextBoxStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            if (option.IsUseFrostedGlassTextBoxStyle)
            {
                int blurDegreeForFrostedGlass = EffectsDesigner.BlurDegreeForFrostedGlassEffect;
                Shape blurImageShape = option.IsUseSpecialEffectStyle
                    ? designer.ApplyBlurEffect(source.SpecialEffectImageFile, blurDegreeForFrostedGlass)
                    : designer.ApplyBlurEffect(degree: blurDegreeForFrostedGlass);
                designer.ApplyFrostedGlassTextBoxEffect(option.FrostedGlassTextBoxColor, option.FrostedGlassTextBoxTransparency,
                    blurImageShape, option.FontSizeIncrease);
                blurImageShape.Delete();
            }
            return new List<Shape>();
        }
    }
}
