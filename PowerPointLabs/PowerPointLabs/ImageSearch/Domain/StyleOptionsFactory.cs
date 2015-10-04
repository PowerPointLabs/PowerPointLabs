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
                    return GetOptionsWithSuitableFontColor();
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return GetOptionsWithSuitableFontColor();
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return GetOptions();
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return GetOptionsWithSuitableFontColor();
                default:
                    return new List<StyleOptions>();
            }
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
