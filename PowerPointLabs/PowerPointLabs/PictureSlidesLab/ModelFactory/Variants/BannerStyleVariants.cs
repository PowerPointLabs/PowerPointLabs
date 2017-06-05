using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    [Export(typeof(IStyleVariants))]
    class BannerStyleVariants : BaseStyleVariants
    {
        public override string GetStyleName()
        {
            return TextCollection.PictureSlidesLabText.StyleNameBanner;
        }

        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new BannerVariantWorker(),
                new BannerTransparencyVariantWorker(),
                new GeneralSpecialEffectsVariantWorker(),
                new BlurVariantWorker(),
                new BrightnessVariantWorker()
            };
        }
    }
}
