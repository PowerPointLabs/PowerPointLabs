using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory
{
    public class StyleVariantsFactory
    {
        public static Dictionary<string, List<StyleVariants>> GetVariants(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.PictureSlidesLabText.StyleNameDirectText:
                    return GetVariantsForDirectText();
                case TextCollection.PictureSlidesLabText.StyleNameBlur:
                    return GetVariantsForBlur();
                case TextCollection.PictureSlidesLabText.StyleNameTextBox:
                    return GetVariantsForTextBox();
                case TextCollection.PictureSlidesLabText.StyleNameBanner:
                    return GetVariantsForBanner();
                case TextCollection.PictureSlidesLabText.StyleNameSpecialEffect:
                    return GetVariantsForSpecialEffect();
                case TextCollection.PictureSlidesLabText.StyleNameOverlay:
                    return GetVariantsForOverlay();
                case TextCollection.PictureSlidesLabText.StyleNameOutline:
                    return GetVariantsForOutline();
                case TextCollection.PictureSlidesLabText.StyleNameFrame:
                    return GetVariantsForFrame();
                case TextCollection.PictureSlidesLabText.StyleNameCircle:
                    return GetVariantsForCircle();
                case TextCollection.PictureSlidesLabText.StyleNameTriangle:
                    return GetVariantsForTriangle();
                default:
                    return new Dictionary<string, List<StyleVariants>>();
            }
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForTriangle()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryTriangleColor, GetTriangleColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTriangleTransparency, GetTriangleTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForCircle()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryCircleColor, GetCircleColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryCircleTransparency, GetCircleTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForFrame()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryFrameColor, GetFrameColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFrameTransparency, GetFrameTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForOutline()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForOverlay()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryOverlayColor, GetOverlayVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryOverlayTransparency, GetOverlayTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForSpecialEffect()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForBanner()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryBannerColor, GetBannerVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBannerTransparency, GetBannerTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForTextBox()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryTextBoxColor, GetTextBoxVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextBoxTransparency, GetTextBoxTransparencyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForBlur()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategorySpecialEffects, GetGeneralSpecialEffectVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForDirectText()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.PictureSlidesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontColor, GetFontColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextGlowColor, GetTextGlowColorVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() },
                { TextCollection.PictureSlidesLabText.VariantCategoryImageReference, GetImageReferenceVariants() }
            };
        }

        private static List<StyleVariants> GetTriangleTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"TriangleTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "5% Transparency"},
                    {"TriangleTransparency", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"TriangleTransparency", 10}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"TriangleTransparency", 15}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"TriangleTransparency", 20}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"TriangleTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"TriangleTransparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"TriangleTransparency", 35}
                })
            };
        }

        private static List<StyleVariants> GetTriangleColorVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"TriangleColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"TriangleColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"TriangleColor", "#FFCC00"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"TriangleColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"TriangleColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"TriangleColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"TriangleColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"TriangleColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetCircleTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"CircleTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "5% Transparency"},
                    {"CircleTransparency", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"CircleTransparency", 10}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"CircleTransparency", 15}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"CircleTransparency", 20}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"CircleTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"CircleTransparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"CircleTransparency", 35}
                })
            };
        }

        private static List<StyleVariants> GetCircleColorVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"CircleColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"CircleColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"CircleColor", "#FFCC00"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"CircleColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"CircleColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"CircleColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"CircleColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"CircleColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetOverlayVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFCC00"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FF0000"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#3DFF8F"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#007FFF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#7800FF"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#001550"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                })
            };
        }

        private static List<StyleVariants> GetOverlayTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 50}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "45% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 45}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 40}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 35}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 20}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"IsUseOverlayStyle", true},
                    {"Transparency", 15}
                })
            };
        }

        private static List<StyleVariants> GetSpecialEffectVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Grayscale"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black and White"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Gotham"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 3}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "HiSatch"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 4}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Invert"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Lomograph"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 6}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Polaroid"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 8}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Sepia"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 9}
                })
            };
        }

        private static List<StyleVariants> GetGeneralSpecialEffectVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Grayscale"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black and White"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Gotham"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 3}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "HiSatch"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 4}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Invert"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Lomograph"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 6}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Polaroid"},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 8}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseSpecialEffectStyle", false},
                    {"SpecialEffect", -1}
                })
            };
        }

        private static List<StyleVariants> GetFrameColorVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"FrameColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"FrameColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"FrameColor", "#FFC500"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"FrameColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"FrameColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"FrameColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"FrameColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"FrameColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetBannerVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FFC500"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetBannerTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 60}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 50}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 40}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 35}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 15}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"IsUseBannerStyle", true},
                    {"BannerTransparency", 0}
                })
            };
        }

        private static List<StyleVariants> GetTextBoxVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#FFC500"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetTextGlowColorVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsUseTextGlow", false},
                    {"TextGlowColor", "#123456"} // match the default init value
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#FFC500"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#7800FF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseTextGlow", true},
                    {"TextGlowColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetTextBoxTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 60}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 50}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 40}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "35% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 35}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "25% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "15% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 15}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxTransparency", 0}
                })
            };
        }

        private static List<StyleVariants> GetBlurVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "100% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 100}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "90% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 95}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "80% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 90}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "70% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 85}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "60% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 80}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 75}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "40% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 70}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Blurriness"},
                    {"IsUseBlurStyle", true},
                    {"BlurDegree", 0}
                })
            };
        }

        private static List<StyleVariants> GetFrameTransparencyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Transparency"},
                    {"FrameTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "10% Transparency"},
                    {"FrameTransparency", 10}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "20% Transparency"},
                    {"FrameTransparency", 20}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "30% Transparency"},
                    {"FrameTransparency", 30}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "40% Transparency"},
                    {"FrameTransparency", 40}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Transparency"},
                    {"FrameTransparency", 50}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "60% Transparency"},
                    {"FrameTransparency", 60}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "70% Transparency"},
                    {"FrameTransparency", 70}
                })
            };
        }

        private static List<StyleVariants> GetBrightnessVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "140% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"Transparency", 60}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "120% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFFFFF"},
                    {"Transparency", 80}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "100% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 100}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "90% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 90}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "80% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 80}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "70% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 70}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "60% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 60}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "50% Brightness"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 50}
                })
            };
        }

        private static List<StyleVariants> GetFontPositionVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Centered"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom-left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 7}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 4}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Original"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Centered-left align"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 5},
                    {"TextBoxAlignment", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom-left align"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8},
                    {"TextBoxAlignment", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Right"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 6}
                })
            };
        }

        private static List<StyleVariants> GetFontColorVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#FFFFFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#000000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#FFD700"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#FF0000"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#3DFF8F"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#007FFF"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#7F00D4"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseTextFormat", true},
                    {"FontColor", "#001550"}
                })
            };
        }

        private static List<StyleVariants> GetFontFamilyVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Segoe UI"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Segoe UI"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Calibri"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Calibri"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Microsoft YaHei"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Microsoft YaHei"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Arial"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Arial"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Courier New"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Courier New"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Trebuchet MS"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Trebuchet MS"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Times New Roman"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Times New Roman"}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Tahoma"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", "Tahoma"}
                })
            };
        }

        private static List<StyleVariants> GetFontSizeIncreaseVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Original Font Size"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +3"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 3}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +6"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 6}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +9"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 9}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +12"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 12}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +15"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 15}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +18"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 18}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +21"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 21}
                })
            };
        }

        private static List<StyleVariants> GetImageReferenceVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "No Effect"},
                    {"IsInsertReference", false},
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Right"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 3},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Left"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 1},
                    {"CitationFontSize", 14},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Right (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 3},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom Left (Small Font)"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 1},
                    {"CitationFontSize", 10},
                    {"ImageReferenceTextBoxColor", ""}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Bottom With Banner"},
                    {"IsInsertReference", true},
                    {"ImageReferenceAlignment", 2},
                    {"CitationFontSize", 12},
                    {"ImageReferenceTextBoxColor", "#000000"} 
                })
            };
        }
    }
}
