using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    [Export(typeof(IStyleVariants))]
    class BlurStyleVariants : BaseStyleVariants
    {
        public override string GetStyleName()
        {
            return PictureSlidesLabText.StyleNameBlur;
        }

        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new BlurVariantWorker(),
                new GeneralSpecialEffectsVariantWorker(),
                new BrightnessVariantWorker()
            };
        }
    }
}
