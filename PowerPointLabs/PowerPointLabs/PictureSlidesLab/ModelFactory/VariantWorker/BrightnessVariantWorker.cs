using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class BrightnessVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryBrightness;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "140% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"OverlayTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "120% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"OverlayTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"OverlayTransparency", 50}
                })
            };
        }
    }
}
