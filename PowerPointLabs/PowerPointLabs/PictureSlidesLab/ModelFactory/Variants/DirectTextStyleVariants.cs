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
        public override string GetStyleName()
        {
            return TextCollection.PictureSlidesLabText.StyleNameDirectText;
        }

        protected override IList<IVariantWorker> GetRequiredVariantWorkers()
        {
            return new List<IVariantWorker>
            {
                new BrightnessVariantWorker()
            };
        }
    }
}
