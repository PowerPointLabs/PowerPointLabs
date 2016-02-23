using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassTextBoxTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryFrostedGlassTextBoxTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Brightness"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 30}
                }),
            };
        }
    }
}
