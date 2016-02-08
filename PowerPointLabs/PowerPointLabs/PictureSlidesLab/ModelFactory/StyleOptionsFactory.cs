using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory
{
    /// <summary>
    /// in order to ensure continuity in the customisation stage,
    /// style option provided from this factory should have corresponding values specified 
    /// in StyleVariantsFactory. e.g., an option generated from this factory has overlay 
    /// transparency of 35, then in order to swap (ensure continuity), it should have a 
    /// variant of overlay transparency of 35. Otherwise it cannot swap and so lose continuity, 
    /// because variants don't match any values in the style option.
    /// </summary>
    public class StyleOptionsFactory
    {
        /// <summary>
        /// get all styles variation options for variation stage usage
        /// </summary>
        /// <returns></returns>
        public static List<List<StyleOption>> GetAllStylesVariationOptions()
        {
            var options = new List<List<StyleOption>>
            {
                GetOptionsForDirectText(),
                GetOptionsForBlur(),
                GetOptionsForTextBox(),
                GetOptionsForBanner(),
                GetOptionsForSpecialEffect(),
                GetOptionsForOverlay(),
                GetOptionsForOutline(),
                GetOptionsForFrame(),
                GetOptionsForCircle(),
                GetOptionsForTriangle(),
            };
            return options;
        }

        /// <summary>
        /// get all styles preview options for preview stage usage
        /// </summary>
        /// <returns></returns>
        public static List<StyleOption> GetAllStylesPreviewOptions()
        {
            var options = new List<StyleOption>
            {
                GetDefaultOptionForDirectText(),
                GetDefaultOptionForBlur(),
                GetDefaultOptionForTextBox(),
                GetDefaultOptionForBanner(),
                GetDefaultOptionForSpecialEffects(),
                GetDefaultOptionForOverlay(),
                GetDefaultOptionForOutline(),
                GetDefaultOptionForFrame(),
                GetDefaultOptionForCircle(),
                GetDefaultOptionForTriangle(),
            };
            return options;
        }

        public static StyleOption GetStylesPreviewOption(string targetStyle)
        {
            var options = GetAllStylesPreviewOptions();
            foreach (var option in options)
            {
                if (option.StyleName == targetStyle)
                {
                    return option;
                }
            }
            return options[0];
        }

        public static List<StyleOption> GetStylesVariationOptions(string targetStyle)
        {
            var allStylesVariationOptions = GetAllStylesVariationOptions();
            foreach (var stylesVariationOptions in allStylesVariationOptions)
            {
                if (stylesVariationOptions[0].StyleName == targetStyle)
                {
                    return stylesVariationOptions;
                }
            }
            return allStylesVariationOptions[0];
        }

        private static List<StyleOption> UpdateStyleName(List<StyleOption> opts, string styleName)
        {
            int i = 0;
            foreach (var styleOption in opts)
            {
                styleOption.StyleName = styleName;
                i++;
            }
            return opts;
        }

        #region Get specific styles preview option

        private static StyleOption GetDefaultOptionForDirectText()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameDirectText,
                TextBoxPosition = 5,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }

        private static StyleOption GetDefaultOptionForBlur()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameBlur,
                IsUseBlurStyle = true,
                BlurDegree = 85,
                TextBoxPosition = 5,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }

        private static StyleOption GetDefaultOptionForTextBox()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameTextBox,
                IsUseTextBoxStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700"
            };
        }

        private static StyleOption GetDefaultOptionForBanner()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameBanner,
                IsUseBannerStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700"
            };
        }

        private static StyleOption GetDefaultOptionForSpecialEffects()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameSpecialEffect,
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }

        private static StyleOption GetDefaultOptionForOverlay()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameOverlay,
                IsUseOverlayStyle = true,
                Transparency = 35,
                OverlayColor = "#007FFF", // blue
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0
            };
        }

        private static StyleOption GetDefaultOptionForOutline()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameOutline,
                IsUseOutlineStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }

        private static StyleOption GetDefaultOptionForFrame()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameFrame,
                IsUseFrameStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }

        private static StyleOption GetDefaultOptionForTriangle()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameTriangle,
                IsUseTriangleStyle = true,
                TriangleColor = "#007FFF", // blue
                TextBoxPosition = 4, // left
                TriangleTransparency = 25
            };
        }

        private static StyleOption GetDefaultOptionForCircle()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameCircle,
                IsUseCircleStyle = true,
                FontColor = "#000000",
                CircleTransparency = 25
            };
        }

        #endregion
        #region Get specific styles variation options

        private static List<StyleOption> GetOptionsForTriangle()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTriangleStyle = true;
                styleOption.TextBoxPosition = 4; // left
                styleOption.TriangleTransparency = 25;
            }
            UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameTriangle);
            return result;
        }

        private static List<StyleOption> GetOptionsForCircle()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseCircleStyle = true;
                styleOption.CircleTransparency = 25;
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameCircle);
        }

        private static List<StyleOption> GetOptionsForFrame()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseFrameStyle = true;
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameFrame);
        }

        private static List<StyleOption> GetOptionsForOutline()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseOutlineStyle = true;
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameOutline);
        }

        private static List<StyleOption> GetOptionsForOverlay()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.Transparency = 35;
            }
            return UpdateStyleName(result,
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
        } 

        private static List<StyleOption> GetOptionsForSpecialEffect()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
        } 

        private static List<StyleOption> GetOptionsForBanner()
        {
            return UpdateStyleName(
                GetOptionsForTextBox(),
                TextCollection.PictureSlidesLabText.StyleNameBanner);
        } 

        private static List<StyleOption> GetOptionsForTextBox()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var option in result)
            {
                option.TextBoxPosition = 7; //bottom-left;
            }
            UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameTextBox);
            return result;
        }

        private static List<StyleOption> GetOptionsForBlur()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameBlur);
        } 

        private static List<StyleOption> GetOptionsForDirectText()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameDirectText);
            return result;
        }

        private static List<StyleOption> GetOptions()
        {
            var result = new List<StyleOption>();
            for (var i = 0; i < 8; i++)
            {
                result.Add(new StyleOption());
            }
            return result;
        }

        private static List<StyleOption> GetOptionsWithSuitableFontColor()
        {
            var result = GetOptions();
            result[0].FontColor = "#000000"; //white(bg color) + black
            result[1].FontColor = "#FFD700"; //black + yellow
            result[2].FontColor = "#000000"; //yellow + black
            result[4].FontColor = "#001550"; //green + dark blue
            result[6].FontColor = "#FFD700"; //purple + yellow
            result[7].FontColor = "#3DFF8F"; //dark blue + green
            return result;
        }
        #endregion
    }
}
