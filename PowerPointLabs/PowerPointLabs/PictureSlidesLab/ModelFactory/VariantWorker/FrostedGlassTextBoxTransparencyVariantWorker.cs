using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassTextBoxTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFrostedGlassTextBoxTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "100% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 100}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "90% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 90}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "80% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 80}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 70}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxTransparency", 30}
                }),
            };
        }
    }
}
