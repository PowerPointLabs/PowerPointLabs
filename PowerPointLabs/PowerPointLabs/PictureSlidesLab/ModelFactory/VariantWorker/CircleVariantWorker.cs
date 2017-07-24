using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class CircleVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryCircleColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"CircleColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"CircleColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"CircleColor", "#FFCC00"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"CircleColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"CircleColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"CircleColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"CircleColor", "#7800FF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"CircleColor", "#001550"}
                })
            };
        }
    }
}
