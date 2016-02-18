using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassBannerTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryFrostedGlassBannerTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Brightness"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerTransparency", 30}
                }),
            };
        }
    }
}
