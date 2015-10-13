using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.Domain
{
    class StyleOptionsFactory
    {
        public static List<StyleOptions> GetOptions(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return GetOptionsForDirectText();
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return GetOptions();
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return GetOptionsForTextBox();
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return GetOptionsForTextBox();
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return GetOptions();
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return GetOptionsWithSuitableFontColor();
                default:
                    return new List<StyleOptions>();
            }
        }

        public static StyleOptions GetDefaultOption(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return GetDefaultOptionForDirectText();
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return GetDefaultOptionForBlur();
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return GetDefaultOptionForTextBox();
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return GetDefaultOptionForBanner();
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return GetDefaultOptionForSpecialEffects();
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return GetDefaultOptionForOverlay();
                default:
                    return new StyleOptions();
            }
        }

        private static StyleOptions GetDefaultOptionForDirectText()
        {
            return new StyleOptions
            {
                TextBoxPosition = 4
            };
        }

        private static StyleOptions GetDefaultOptionForBlur()
        {
            return new StyleOptions
            {
                IsUseBlurStyle = true,
                TextBoxPosition = 4
            };
        }

        private static StyleOptions GetDefaultOptionForTextBox()
        {
            return new StyleOptions
            {
                IsUseTextBoxStyle = true,
                TextBoxPosition = 7
            };
        }

        private static StyleOptions GetDefaultOptionForBanner()
        {
            return new StyleOptions
            {
                IsUseBannerStyle = true,
                TextBoxPosition = 7
            };
        }

        private static StyleOptions GetDefaultOptionForSpecialEffects()
        {
            return new StyleOptions
            {
                IsUseSpecialEffectStyle = true
            };
        }

        private static StyleOptions GetDefaultOptionForOverlay()
        {
            return new StyleOptions
            {
                IsUseOverlayStyle = true
            };
        }

        private static List<StyleOptions> GetOptionsForTextBox()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var option in result)
            {
                option.TextBoxPosition = 7;//bottom-left;
            }
            return result;
        }

        private static List<StyleOptions> GetOptionsForDirectText()
        {
            var result = GetOptions();
            result[0].FontColor = "#000000";
            result[1].FontColor = "#000000";
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
            return result;
        } 
    }
}
