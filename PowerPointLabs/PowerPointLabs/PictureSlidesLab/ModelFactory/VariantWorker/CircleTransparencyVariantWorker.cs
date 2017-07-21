using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class CircleTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryCircleTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"CircleTransparency", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "5% Transparency"},
                    {"CircleTransparency", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"CircleTransparency", 10}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"CircleTransparency", 15}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"CircleTransparency", 20}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"CircleTransparency", 25}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"CircleTransparency", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"CircleTransparency", 35}
                })
            };
        }
    }
}
