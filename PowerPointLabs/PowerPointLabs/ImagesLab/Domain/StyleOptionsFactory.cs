using System.Collections.Generic;

namespace PowerPointLabs.ImagesLab.Domain
{
    class StyleOptionsFactory
    {
        public static List<StyleOptions> GetOptions(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return UpdateStyleName(
                        GetOptionsForDirectText(),
                        TextCollection.ImagesLabText.StyleNameDirectText);
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return UpdateStyleName(
                        GetOptions(),
                        TextCollection.ImagesLabText.StyleNameBlur);
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return UpdateStyleName(
                        GetOptionsForTextBox(),
                        TextCollection.ImagesLabText.StyleNameTextBox);
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return UpdateStyleName(
                        GetOptionsForTextBox(),
                        TextCollection.ImagesLabText.StyleNameBanner);
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return UpdateStyleName(
                        GetOptions(),
                        TextCollection.ImagesLabText.StyleNameSpecialEffect);
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return UpdateStyleName(
                        GetOptionsWithSuitableFontColor(),
                        TextCollection.ImagesLabText.StyleNameOverlay);
                default:
                    return new List<StyleOptions>();
            }
        }

        public static StyleOptions GetDefaultOption(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return UpdateStyleName(
                        GetDefaultOptionForDirectText(),
                        TextCollection.ImagesLabText.StyleNameDirectText);
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return UpdateStyleName(
                        GetDefaultOptionForBlur(),
                        TextCollection.ImagesLabText.StyleNameBlur);
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return UpdateStyleName(
                        GetDefaultOptionForTextBox(),
                        TextCollection.ImagesLabText.StyleNameTextBox);
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return UpdateStyleName(
                        GetDefaultOptionForBanner(),
                        TextCollection.ImagesLabText.StyleNameBanner);
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return UpdateStyleName(
                        GetDefaultOptionForSpecialEffects(),
                        TextCollection.ImagesLabText.StyleNameSpecialEffect);
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return UpdateStyleName(
                        GetDefaultOptionForOverlay(),
                        TextCollection.ImagesLabText.StyleNameOverlay);
                default:
                    return new StyleOptions();
            }
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

        private static StyleOptions UpdateStyleName(StyleOptions opt, string styleName)
        {
            opt.StyleName = styleName;
            return opt;
        }

        private static StyleOptions GetDefaultOptionForDirectText()
        {
            return new StyleOptions
            {
                TextBoxPosition = 5
            };
        }

        private static StyleOptions GetDefaultOptionForBlur()
        {
            return new StyleOptions
            {
                IsUseBlurStyle = true,
                BlurDegree = 85,
                TextBoxPosition = 5
            };
        }

        private static StyleOptions GetDefaultOptionForTextBox()
        {
            return new StyleOptions
            {
                IsUseTextBoxStyle = true,
                TextBoxPosition = 7,
                TextBoxOverlayColor = "#000000",
                FontColor = "#FFD700"
            };
        }

        private static StyleOptions GetDefaultOptionForBanner()
        {
            return new StyleOptions
            {
                IsUseBannerStyle = true,
                TextBoxPosition = 7,
                TextBoxOverlayColor = "#000000",
                FontColor = "#FFD700"
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
                IsUseOverlayStyle = true,
                Transparency = 35,
                OverlayColor = "#007FFF"
            };
        }

        private static List<StyleOptions> GetOptionsForTextBox()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var option in result)
            {
                option.TextBoxPosition = 7;//bottom-left;
                option.Transparency = 100;
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
            for (var i = 0; i < 8; i++)
            {
                result[i].Transparency = 35;
            }
            return result;
        } 
    }
}
