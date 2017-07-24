using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrostedGlassTextBoxVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFrostedGlassTextBoxColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#FFC500"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#7800FF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseFrostedGlassTextBoxStyle", true},
                    {"FrostedGlassTextBoxColor", "#001550"}
                })
            };
        }
    }
}
