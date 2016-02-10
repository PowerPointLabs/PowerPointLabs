using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 2)]
    class FontPositionVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryTextPosition;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Centered"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom-left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 7}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 4}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Original"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Centered-left align"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5},
                    {"TextBoxAlignment", 1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom-left align"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8},
                    {"TextBoxAlignment", 1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Right"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 6}
                })
            };
        }
    }
}
