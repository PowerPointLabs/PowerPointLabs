using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    [Export("GeneralVariantWorker", typeof(IVariantWorker))]
    [ExportMetadata("GeneralVariantWorkerOrder", 5)]
    class TextPositionVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryTextPosition;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Centered"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom-left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 7},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 4},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Centered-left align"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5},
                    {"TextBoxAlignment", 1},
                    {"BannerDirection", 1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Top"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 2},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Right"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 6},
                    {"TextBoxAlignment", 0},
                    {"BannerDirection", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 10}, // No effect. see StyleOption.Text.cs
                    {"TextBoxAlignment", 4}, // No effect. see StyleOption.Text.cs
                    {"BannerDirection", 0}
                }),
            };
        }
    }
}
