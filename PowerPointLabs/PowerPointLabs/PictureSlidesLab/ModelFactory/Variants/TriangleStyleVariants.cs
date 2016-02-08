using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    class TriangleStyleVariants : BaseStyleVariants
    {
        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new TriangleVariantWorker(),
                new TriangleTransparencyVariantWorker(),
                new GeneralSpecialEffectsVariantWorker(),
                new BlurVariantWorker(),
                new BrightnessVariantWorker(),
                new FontColorVariantWorker(),
                new TextGlowVariantWorker(),
                new FontFamilyVariantWorker(),
                new FontSizeIncreaseVariantWorker(),
                new PictureCitationVariantWorker()
            };
        }

        public override string GetStyleName()
        {
            return TextCollection.PictureSlidesLabText.StyleNameTriangle;
        }
    }
}
