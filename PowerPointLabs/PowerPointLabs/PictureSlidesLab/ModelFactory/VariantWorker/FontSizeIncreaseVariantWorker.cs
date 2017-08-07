using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 2)]
    class FontSizeIncreaseVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFontSizeIncrease;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +0"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +5"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +10"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 10}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +20"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 20}
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
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", -1}
                })
            };
        }
    }
}
