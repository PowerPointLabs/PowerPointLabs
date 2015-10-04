using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.Domain
{
    class StyleVariantsFactory
    {
        public static Dictionary<string, List<StyleVariants>> GetVariants(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return GetVariantsForDirectText();
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return GetVariantsForBlur();
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return GetVariantsForTextBox();
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return GetVariantsForBanner();
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return GetVariantsForSpecialEffect();
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return GetVariantsForOverlay();
                default:
                    return new Dictionary<string, List<StyleVariants>>();
            }
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForOverlay()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategoryOverlayColor, GetOverlayVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForSpecialEffect()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategorySpecialEffects, GetSpecialEffectVariants() },
                { TextCollection.ImagesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForBanner()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategoryBannerColor, GetBannerVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForTextBox()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategoryTextBoxColor, GetTextBoxVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForBlur()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategoryBlurriness, GetBlurVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
            };
        }

        private static Dictionary<string, List<StyleVariants>> GetVariantsForDirectText()
        {
            return new Dictionary<string, List<StyleVariants>>
            {
                { TextCollection.ImagesLabText.VariantCategoryBrightness, GetBrightnessVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextColor, GetFontColorVariants() },
                { TextCollection.ImagesLabText.VariantCategoryTextPosition, GetFontPositionVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontFamily, GetFontFamilyVariants() },
                { TextCollection.ImagesLabText.VariantCategoryFontSizeIncrease, GetFontSizeIncreaseVariants() }
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
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#000000"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FFCC00"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#FF0000"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#3DFF8F"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#007FFF"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#7800FF"},
                    {"Transparency", 40},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseOverlayStyle", true},
                    {"OverlayColor", "#001550"},
                    {"Transparency", 25},
                    {"IsUseSpecialEffectStyle", true},
                    {"SpecialEffect", 0}
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

        private static List<StyleVariants> GetBannerVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "White"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#FFFFFF"},
                    {"BannerTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#000000"},
                    {"BannerTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#FFC500"},
                    {"BannerTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#FF0000"},
                    {"BannerTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#3DFF8F"},
                    {"BannerTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#007FFF"},
                    {"BannerTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#7800FF"},
                    {"BannerTransparency", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseBannerStyle", true},
                    {"BannerOverlayColor", "#001550"},
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
                    {"TextBoxOverlayColor", "#FFFFFF"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Black"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#000000"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Yellow"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#FFC500"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Red"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#FF0000"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Green"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#3DFF8F"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Blue"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#007FFF"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Purple"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#7800FF"},
                    {"TextBoxTransparency", 25}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Dark Blue"},
                    {"IsUseTextBoxStyle", true},
                    {"TextBoxOverlayColor", "#001550"},
                    {"TextBoxTransparency", 25}
                })
            };
        }

        private static List<StyleVariants> GetBlurVariants()
        {
            return new List<StyleVariants>
            {
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "0% Blurriness"},
                    {"IsUseBlurStyle", false}
                }),
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
                    {"OptionName", "Bottom"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 8}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Top-left"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Top"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 2}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Top-right"},
                    {"IsUseTextFormat", true},
                    {"TextBoxPosition", 3}
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
                    {"OptionName", "Original Font"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 5}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Segoe UI"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 0}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Segoe UI Light"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 1}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Calibri"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 2}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Calibri Light"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 3}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Trebuchet MS"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 4}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Times New Roman"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 6}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Tahoma"},
                    {"IsUseTextFormat", true},
                    {"FontFamily", 7}
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
                    {"OptionName", "Font Size +2"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 2}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +4"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 4}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +6"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 6}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +8"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 8}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +10"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 10}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +12"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 12}
                }),
                new StyleVariants(new Dictionary<string, object>
                {
                    {"OptionName", "Font Size +14"},
                    {"IsUseTextFormat", true},
                    {"FontSizeIncrease", 14}
                })
            };
        }
    }
}
