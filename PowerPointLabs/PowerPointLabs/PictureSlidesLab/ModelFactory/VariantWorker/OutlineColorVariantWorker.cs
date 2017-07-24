using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class OutlineColorVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryOutlineColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#FFD700"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#7F00D4"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseOutlineStyle", true},
                    {"OutlineColor", "#001550"}
                })
            };
        }
    }
}
