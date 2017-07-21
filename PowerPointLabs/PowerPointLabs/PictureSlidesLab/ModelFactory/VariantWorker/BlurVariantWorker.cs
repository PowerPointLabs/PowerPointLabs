using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class BlurVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryBlurriness;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 95}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 85}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 75}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "0% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 0}
                })
            };
        }
    }
}
