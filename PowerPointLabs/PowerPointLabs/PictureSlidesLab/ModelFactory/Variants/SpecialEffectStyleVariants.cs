using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    [Export(typeof(IStyleVariants))]
    class SpecialEffectStyleVariants : BaseStyleVariants
    {
        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new SpecialEffectsVariantWorker(),
                new BlurVariantWorker(),
                new BrightnessVariantWorker(),
                new FontColorVariantWorker(),
                new TextGlowVariantWorker(),
                new FontPositionVariantWorker(),
                new FontFamilyVariantWorker(),
                new FontSizeIncreaseVariantWorker(),
                new PictureCitationVariantWorker()
            };
        }

        public override string GetStyleName()
        {
            return TextCollection.PictureSlidesLabText.StyleNameSpecialEffect;
        }
    }
}
