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
                    {"OptionName", "No Effect"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", -1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +3"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 3}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +10"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 10}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +15"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 15}
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
                })
            };
        }
    }
}
