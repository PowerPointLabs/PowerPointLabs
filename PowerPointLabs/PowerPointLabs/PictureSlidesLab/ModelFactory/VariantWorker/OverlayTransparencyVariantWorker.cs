using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class OverlayTransparencyVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return TextCollection.PictureSlidesLabText.VariantCategoryOverlayTransparency;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 50}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "45% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 45}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 40}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 35}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 30}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 25}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 20}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 15}
                })
            };
        }
    }
}
