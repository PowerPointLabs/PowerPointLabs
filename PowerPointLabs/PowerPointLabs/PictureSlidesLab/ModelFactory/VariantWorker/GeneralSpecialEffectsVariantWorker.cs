using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker
{
    class GeneralSpecialEffectsVariantWorker : IVariantWorker
    {
        public string GetVariantName()
        {
            return PictureSlidesLabText.VariantCategorySpecialEffects;
        }

        public List<StyleVariant> GetVariants()
        {
            return new List<StyleVariant>
            {
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Grayscale"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Black and White"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 1}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Gotham"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 3}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "HiSatch"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 4}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Invert"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 5}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Lomograph"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 6}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "Polaroid"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 8}
                }),
                new StyleVariant(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseSpecialEffectStyle", false},
                    {"SpecialEffect", -1}
                })
            };
        }
    }
}
