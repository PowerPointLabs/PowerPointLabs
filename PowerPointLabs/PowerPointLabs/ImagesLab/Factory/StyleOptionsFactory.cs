using System.Collections.Generic;
using PowerPointLabs.ImagesLab.Domain;

namespace PowerPointLabs.ImagesLab.Factory
{
    class StyleOptionsFactory
    {
        /// <summary>
        /// get all styles variation options for variation stage usage
        /// </summary>
        /// <returns></returns>
        public static List<List<StyleOptions>> GetAllStylesVariationOptions()
        {
            var options = new List<List<StyleOptions>>
            {
                GetOptionsForDirectText(),
                GetOptionsForBlur(),
                GetOptionsForTextBox(),
                GetOptionsForBanner(),
                GetOptionsForSpecialEffect(),
                GetOptionsForOverlay(),
                GetOptionsForOutline(),
                GetOptionsForFrame(),
            };
            return options;
        }

        /// <summary>
        /// get all styles preview options for preview stage usage
        /// </summary>
        /// <returns></returns>
        public static List<StyleOptions> GetAllStylesPreviewOptions()
        {
            var options = new List<StyleOptions>
            {
                GetDefaultOptionForDirectText(),
                GetDefaultOptionForBlur(),
                GetDefaultOptionForTextBox(),
                GetDefaultOptionForBanner(),
                GetDefaultOptionForSpecialEffects(),
                GetDefaultOptionForOverlay(),
                GetDefaultOptionForOutline(),
                GetDefaultOptionForFrame(),
            };
            return options;
        }

        public static StyleOptions GetStylesPreviewOption(string targetStyle)
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

        public static List<StyleOptions> GetStylesVariationOptions(string targetStyle)
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

        private static List<StyleOptions> UpdateStyleName(List<StyleOptions> opts, string styleName)
        {
            int i = 0;
            foreach (var styleOption in opts)
            {
                styleOption.StyleName = styleName;
                styleOption.VariantIndex = i;
                i++;
            }
            return opts;
        }

        #region Get specific styles preview option

        private static StyleOptions GetDefaultOptionForDirectText()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameDirectText,
                TextBoxPosition = 5
            };
        }

        private static StyleOptions GetDefaultOptionForBlur()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameBlur,
                IsUseBlurStyle = true,
                BlurDegree = 85,
                TextBoxPosition = 5
            };
        }

        private static StyleOptions GetDefaultOptionForTextBox()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameTextBox,
                IsUseTextBoxStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700"
            };
        }

        private static StyleOptions GetDefaultOptionForBanner()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameBanner,
                IsUseBannerStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700"
            };
        }

        private static StyleOptions GetDefaultOptionForSpecialEffects()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameSpecialEffect,
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0
            };
        }

        private static StyleOptions GetDefaultOptionForOverlay()
        {
            return new StyleOptions
            {
                StyleName = TextCollection.ImagesLabText.StyleNameOverlay,
                IsUseOverlayStyle = true,
                Transparency = 35,
                OverlayColor = "#007FFF"
            };
        }

        private static StyleOptions GetDefaultOptionForOutline()
        {
            return new StyleOptions()
            {
                StyleName = TextCollection.ImagesLabText.StyleNameOutline,
                IsUseOutlineStyle = true
            };
        }

        private static StyleOptions GetDefaultOptionForFrame()
        {
            return new StyleOptions()
            {
                StyleName = TextCollection.ImagesLabText.StyleNameFrame,
                IsUseFrameStyle = true
            };
        }

        #endregion
        #region Get specific styles variation options

        private static List<StyleOptions> GetOptionsForFrame()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseFrameStyle = true;
            }
            UpdateStyleName(
                result,
                TextCollection.ImagesLabText.StyleNameFrame);
            return result;
        }

        private static List<StyleOptions> GetOptionsForOutline()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseOutlineStyle = true;
            }
            UpdateStyleName(
                result,
                TextCollection.ImagesLabText.StyleNameOutline);
            return result;
        }

        private static List<StyleOptions> GetOptionsForOverlay()
        {
            return UpdateStyleName(
                GetOptionsWithSuitableFontColor(),
                TextCollection.ImagesLabText.StyleNameOverlay);
        } 

        private static List<StyleOptions> GetOptionsForSpecialEffect()
        {
            return UpdateStyleName(
                GetOptions(),
                TextCollection.ImagesLabText.StyleNameSpecialEffect);
        } 

        private static List<StyleOptions> GetOptionsForBanner()
        {
            return UpdateStyleName(
                GetOptionsForTextBox(),
                TextCollection.ImagesLabText.StyleNameBanner);
        } 

        private static List<StyleOptions> GetOptionsForTextBox()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var option in result)
            {
                option.TextBoxPosition = 7;//bottom-left;
                option.Transparency = 100;
            }
            UpdateStyleName(
                result,
                TextCollection.ImagesLabText.StyleNameTextBox);
            return result;
        }

        private static List<StyleOptions> GetOptionsForBlur()
        {
            return UpdateStyleName(
                GetOptions(),
                TextCollection.ImagesLabText.StyleNameBlur);
        } 

        private static List<StyleOptions> GetOptionsForDirectText()
        {
            var result = GetOptions();
            result[0].FontColor = "#000000";
            result[1].FontColor = "#000000";
            UpdateStyleName(
                result,
                TextCollection.ImagesLabText.StyleNameDirectText);
            return result;
        }

        private static List<StyleOptions> GetOptions()
        {
            var result = new List<StyleOptions>();
            for (var i = 0; i < 8; i++)
            {
                result.Add(new StyleOptions());
            }
            return result;
        }

        private static List<StyleOptions> GetOptionsWithSuitableFontColor()
        {
            var result = GetOptions();
            result[0].FontColor = "#000000";//white(bg color) + black
            result[1].FontColor = "#FFD700";//black + yellow
            result[2].FontColor = "#000000";//yellow + black
            result[4].FontColor = "#001550";//green + dark blue
            result[6].FontColor = "#FFD700";//purple + yellow
            result[7].FontColor = "#3DFF8F";//dark blue + green
            for (var i = 0; i < 8; i++)
            {
                result[i].Transparency = 35;
            }
            return result;
        }
        #endregion
    }
}
