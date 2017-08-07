using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassBannerVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFrostedGlassBannerColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#FFC500"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#7800FF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseFrostedGlassBannerStyle", true},
                    {"FrostedGlassBannerColor", "#001550"}
                })
            };
        }
    }
}
