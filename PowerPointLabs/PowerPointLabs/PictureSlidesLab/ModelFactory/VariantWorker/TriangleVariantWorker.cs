using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class TriangleVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryTriangleColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"TriangleColor", "#FFFFFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"TriangleColor", "#000000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"TriangleColor", "#FFCC00"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"TriangleColor", "#FF0000"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"TriangleColor", "#3DFF8F"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"TriangleColor", "#007FFF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"TriangleColor", "#7800FF"}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"TriangleColor", "#001550"}
                })
            };
        }
    }
}
