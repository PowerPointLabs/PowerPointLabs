using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FontSizeIncreaseVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Original Font Size"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +3"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 3}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +6"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 6}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +9"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 9}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +12"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 12}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +15"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 15}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +18"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 18}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +21"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 21}
                })
            };
        }
    }
}
