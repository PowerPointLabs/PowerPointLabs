using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 10)]
    class FrostedGlassBannerStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape)
        {
            var result = new List<Shape>();
            if (option.IsUseFrostedGlassBannerStyle)
            {
                var blurDegreeForFrostedGlass = EffectsDesigner.BlurDegreeForFrostedGlassEffect;
                var blurImageShape = option.IsUseSpecialEffectStyle
                    ? designer.ApplyBlurEffect(source.SpecialEffectImageFile, blurDegreeForFrostedGlass)
                    : designer.ApplyBlurEffect(degree: blurDegreeForFrostedGlass);
                var banner = designer.ApplyFrostedGlassBannerEffect(option.GetBannerDirection(), option.GetTextBoxPosition(),
                    blurImageShape, option.FrostedGlassBannerColor, option.FrostedGlassBannerTransparency);
                result.Add(banner);
                blurImageShape.Delete();
            }
            return result;
        }
    }
}
