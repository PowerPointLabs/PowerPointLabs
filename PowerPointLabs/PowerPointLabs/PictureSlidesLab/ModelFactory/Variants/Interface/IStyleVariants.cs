using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface
{
    interface IStyleVariants
    {
        Dictionary<string, List<StyleVariant>> GetVariantsForStyle();
    }
}
