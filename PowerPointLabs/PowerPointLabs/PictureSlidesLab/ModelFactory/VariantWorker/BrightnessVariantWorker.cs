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
                    {"Transparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "120% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"Transparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 50}
                })
            };
        }
    }
}
