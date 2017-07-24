using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class OverlayVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategoryOverlayColor;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFCC00"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FF0000"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#3DFF8F"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#007FFF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#7800FF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#001550"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                })
            };
        }
    }
}
