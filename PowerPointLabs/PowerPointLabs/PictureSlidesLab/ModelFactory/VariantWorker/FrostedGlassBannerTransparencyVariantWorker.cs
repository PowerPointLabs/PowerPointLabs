using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassBannerTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFrostedGlassBannerTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 30}
                }),
            };
        }
    }
}
