using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 0)]
    class FontFamilyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFontFamily;
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
                    {"OptionName", "Century Gothic"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Century Gothic"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Times New Roman Italic"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Times New Roman Italic"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Segoe Print"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Segoe Print"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Courier New"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Courier New"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Times New Roman"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Times New Roman"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Impact"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Impact"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", ""}
                }),
            };
        }
    }
}
