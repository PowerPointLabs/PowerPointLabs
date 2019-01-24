using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;

using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Factory
{
    [Export(typeof(StyleWorkerFactory))]
    class StyleWorkerFactory
    {
        public IEnumerable<IStyleWorker> StyleWorkers
        {
            get
            {
                return ImportedStyleWorkers
                    .OrderBy(worker => worker.Metadata.WorkerOrder)
                    .Select(worker => worker.Value);
            }
        }

        [ImportMany(typeof(IStyleWorker))]
        private IEnumerable<Lazy<IStyleWorker, IWorkerOrderMetadata>> ImportedStyleWorkers { get; set; }
    }
}
