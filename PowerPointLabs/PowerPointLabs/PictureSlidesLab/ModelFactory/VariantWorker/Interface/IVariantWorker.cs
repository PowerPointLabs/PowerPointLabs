using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface
{
    interface IVariantWorker
    {
        string GetVariantName();

        List<StyleVariant> GetVariants();
    }
}
