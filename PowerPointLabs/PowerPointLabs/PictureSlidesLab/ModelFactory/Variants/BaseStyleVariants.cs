using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants
{
    abstract class BaseStyleVariants : IStyleVariants
    {
        [ImportMany("GeneralVariantWorker", typeof(IVariantWorker))]
        private IEnumerable<Lazy<IVariantWorker, IGeneralVariantWorkerOrderMetadata>> ImportedGeneralVariantWorkers { get; set; }

        public abstract string GetStyleName();

        public Dictionary<string, List<StyleVariant>> GetVariantsForStyle()
        {
            IList<IVariantWorker> workers = GetRequiredVariantWorkers();
            IEnumerable<IVariantWorker> orderedGeneralVariantWorkers = ImportedGeneralVariantWorkers
                .OrderBy(worker => worker.Metadata.GeneralVariantWorkerOrder)
                .Select(worker => worker.Value);
            foreach (IVariantWorker importedGeneralVariantWorker in orderedGeneralVariantWorkers)
            {
                workers.Add(importedGeneralVariantWorker);
            }
            return workers.ToDictionary(
                worker => worker.GetVariantName(), 
                worker => worker.GetVariants());
        }

        protected abstract IList<IVariantWorker> GetRequiredVariantWorkers();
    }
}
