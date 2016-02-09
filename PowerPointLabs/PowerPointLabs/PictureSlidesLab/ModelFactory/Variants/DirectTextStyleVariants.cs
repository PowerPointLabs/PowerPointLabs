using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    [Export(typeof(IStyleVariants))]
    class DirectTextStyleVariants : BaseStyleVariants
    {
        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
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
            return TextCollection.PictureSlidesLabText.StyleNameDirectText;
        }
    }
}
