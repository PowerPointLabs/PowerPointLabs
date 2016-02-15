using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 4)]
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
                    {"OptionName", "Font Size +6"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 6}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +12"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 12}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +18"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 18}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +24"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 24}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +30"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +45"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 45}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +65"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 65}
                })
            };
        }
    }
}
