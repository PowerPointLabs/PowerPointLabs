using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FontFamilyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryFontFamily;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Segoe UI"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Segoe UI"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Calibri"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Calibri"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Microsoft YaHei"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Microsoft YaHei"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Arial"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Arial"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Courier New"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Courier New"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Trebuchet MS"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Trebuchet MS"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Times New Roman"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Times New Roman"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Tahoma"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Tahoma"}
                })
            };
        }
    }
}
