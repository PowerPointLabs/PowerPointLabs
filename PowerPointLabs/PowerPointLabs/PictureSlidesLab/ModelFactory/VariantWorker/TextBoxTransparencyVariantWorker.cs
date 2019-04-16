using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class TextBoxTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryTextBoxTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 35}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 15}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 0}
                })
            };
        }
    }
}
