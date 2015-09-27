using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.Domain
{
    class StyleOptionsFactory
    {
        public static IList<StyleOptions> GetVariationOptions(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return GetOptionsForDirectText();
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return GetOptionsForBlur();
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return GetOptionsForTextBox();
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return GetOptionsForBanner();
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return GetOptionsForSpecialEffect();
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return GetOptionsForOverlay();
                default:
                    return new List<StyleOptions>();
            }
        }

        private static IList<StyleOptions> GetOptionsForOverlay()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "Black & Yellow",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFD700",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 50,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Black & Red",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FF0000",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 50,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Dark Blue & Light Green",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#3DFF8F",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#001550",
                    Transparency = 50,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Purple & Yellow",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFD700",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#4E0090",
                    Transparency = 50,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Blue",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFFFFF",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#007FFF",
                    Transparency = 60,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Purple",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFFFFF",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#7F00D4",
                    Transparency = 60,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Brown",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFFFFF",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#FFCC00",
                    Transparency = 60,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "Red",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#FFFFFF",
                    IsUseOverlayStyle = true,
                    OverlayColor = "#DD0000",
                    Transparency = 60,
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                }
            };
        } 

        private static IList<StyleOptions> GetOptionsForSpecialEffect()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "Grayscale",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 0
                },
                new StyleOptions
                {
                    OptionName = "BlackAndWhite",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#000000",
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 1
                },
                new StyleOptions
                {
                    OptionName = "HiSatch",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 4
                },
                new StyleOptions
                {
                    OptionName = "Lomograph",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    FontColor = "#000000",
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 6
                },
                new StyleOptions
                {
                    OptionName = "HiSatch and Blur",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 4,
                    IsUseBlurStyle = true,
                    BlurDegree = 99
                },
                new StyleOptions
                {
                    OptionName = "Sepia",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseSpecialEffectStyle = true,
                    SpecialEffect = 9
                }
            };
        } 

        private static IList<StyleOptions> GetOptionsForBanner()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "Bottom-left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#000000",
                    BannerTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Centered",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#000000",
                    BannerTransparency = 25,
                    BannerDirection = 1
                },
                new StyleOptions
                {
                    OptionName = "Left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#000000",
                    BannerTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Red",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#D84825",
                    BannerTransparency = 0
                },
                new StyleOptions
                {
                    OptionName = "Blue",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#007FFF",
                    BannerTransparency = 0
                },
                new StyleOptions
                {
                    OptionName = "Purple",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#7800FF",
                    BannerTransparency = 0
                },
                new StyleOptions
                {
                    OptionName = "Black",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    FontColor = "#FDCB00",
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#000000",
                    BannerTransparency = 0
                },
                new StyleOptions
                {
                    OptionName = "White",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    FontColor = "#000000",
                    IsUseBannerStyle = true,
                    BannerOverlayColor = "#FFFFFF",
                    BannerTransparency = 25
                },
            };
        } 

        private static IList<StyleOptions> GetOptionsForTextBox()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "Left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#000000",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Centered",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#000000",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Bottom-left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#000000",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Bottom",
                    IsUseTextFormat = true,
                    TextBoxPosition = 8,//Bottom
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#000000",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Blue",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#007FFF",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Purple",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#7800FF",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "White",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    FontColor = "#000000",
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#FFFFFF",
                    TextBoxTransparency = 25
                },
                new StyleOptions
                {
                    OptionName = "Yellow",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom-left
                    FontColor = "#000000",
                    IsUseTextBoxStyle = true,
                    TextBoxOverlayColor = "#FFC500",
                    TextBoxTransparency = 25
                },
            };
        } 

        private static IList<StyleOptions> GetOptionsForBlur()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "100% Blurness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBlurStyle = true,
                    BlurDegree = 100
                },
                new StyleOptions
                {
                    OptionName = "80% Blurness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBlurStyle = true,
                    BlurDegree = 95
                },
                new StyleOptions
                {
                    OptionName = "60% Blurness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBlurStyle = true,
                    BlurDegree = 90
                },
                new StyleOptions
                {
                    OptionName = "40% Blurness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBlurStyle = true,
                    BlurDegree = 85
                },
                new StyleOptions
                {
                    OptionName = "20% Blurness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseBlurStyle = true,
                    BlurDegree = 80
                },
                new StyleOptions
                {
                    OptionName = "60% Blurness, Centered",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseBlurStyle = true,
                    BlurDegree = 90
                },
                new StyleOptions
                {
                    OptionName = "60% Blurness, Bottom-left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom left
                    IsUseBlurStyle = true,
                    BlurDegree = 90
                },
                new StyleOptions
                {
                    OptionName = "60% Blurness, Bottom",
                    IsUseTextFormat = true,
                    TextBoxPosition = 8,//Bottom
                    IsUseBlurStyle = true,
                    BlurDegree = 90
                }
            };
        } 

        private static IList<StyleOptions> GetOptionsForDirectText()
        {
            return new List<StyleOptions>
            {
                new StyleOptions
                {
                    OptionName = "120% Brightness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#FFFFFF",
                    Transparency = 80
                },
                new StyleOptions
                {
                    OptionName = "100% Brightness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 100
                },
                new StyleOptions
                {
                    OptionName = "80% Brightness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 80
                },
                new StyleOptions
                {
                    OptionName = "60% Brightness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 60
                },
                new StyleOptions
                {
                    OptionName = "40% Brightness",
                    IsUseTextFormat = true,
                    TextBoxPosition = 4,//Left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 40
                },
                new StyleOptions
                {
                    OptionName = "Centered",
                    IsUseTextFormat = true,
                    TextBoxPosition = 5,//Centered
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 80
                },
                new StyleOptions
                {
                    OptionName = "Bottom-left",
                    IsUseTextFormat = true,
                    TextBoxPosition = 7,//Bottom left
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 80
                },
                new StyleOptions
                {
                    OptionName = "Bottom",
                    IsUseTextFormat = true,
                    TextBoxPosition = 8,//Bottom 
                    IsUseOverlayStyle = true,
                    OverlayColor = "#000000",
                    Transparency = 80
                }
            };
        }
    }
}
