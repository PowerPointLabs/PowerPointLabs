using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class TriangleTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryTriangleTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"TriangleTransparency", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "5% Transparency"},
                    {"TriangleTransparency", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"TriangleTransparency", 10}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"TriangleTransparency", 15}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"TriangleTransparency", 20}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"TriangleTransparency", 25}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"TriangleTransparency", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"TriangleTransparency", 35}
                })
            };
        }
    }
}
