using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class BannerVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryBannerColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FFC500"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#7800FF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#001550"}
                })
            };
        }
    }
}
