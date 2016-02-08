using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    class OutlineStyleVariants : BaseStyleVariants
    {
        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new FontColorVariantWorker(),
                new GeneralSpecialEffectsVariantWorker(),
                new BlurVariantWorker(),
                new BrightnessVariantWorker(),
                new TextGlowVariantWorker(),
                new FontPositionVariantWorker(),
                new FontFamilyVariantWorker(),
                new FontSizeIncreaseVariantWorker(),
                new PictureCitationVariantWorker()
            };
        }

        public override string GetStyleName()
        {
            return TextCollection.PictureSlidesLabText.StyleNameOutline;
        }
    }
}
