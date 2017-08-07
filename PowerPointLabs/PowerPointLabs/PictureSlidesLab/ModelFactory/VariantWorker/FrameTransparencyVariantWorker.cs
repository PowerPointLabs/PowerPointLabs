using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class FrameTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryFrameTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"FrameTransparency", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"FrameTransparency", 10}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"FrameTransparency", 20}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"FrameTransparency", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"FrameTransparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"FrameTransparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"FrameTransparency", 60}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "70% Transparency"},
                    {"FrameTransparency", 70}
                })
            };
        }
    }
}
