using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    abstract class BaseStyleVariants : IStyleVariants
    {
        protected abstract IList<IVariantWorker> GetRequiredVariantWorkers();

        public Dictionary<string, List<StyleVariant>> GetVariantsForStyle()
        {
            var workers = GetRequiredVariantWorkers();
            return workers.ToDictionary(
                worker => worker.GetVariantName(), 
                worker => worker.GetVariants());
        }
    }
}
