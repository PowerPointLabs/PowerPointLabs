using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerPointLabs.ColorThemes
{
    public class ThemeResourceDictionary : ResourceDictionary
    {
        public static readonly string PathToThemesFolder = "pack://application:,,,/PowerPointLabs;component/Resources/Themes/";

        public static ThemeResourceDictionary FromColorTheme(ColorTheme colorTheme)
        {
            string themeName = "";
            switch (colorTheme.ThemeId)
            {
                case ColorTheme.BLACK:
                    themeName = "Black";
                    break;
                case ColorTheme.COLORFUL:
                    themeName = "Colorful";
                    break;
                case ColorTheme.DARK_GRAY:
                    themeName = "DarkGray";
                    break;
                case ColorTheme.WHITE:
                    themeName = "White";
                    break;
                default:
                    themeName = "Colorful";
                    break;
            }

            return new ThemeResourceDictionary
            {
                Source = new Uri(PathToThemesFolder + themeName + "Theme.xaml", UriKind.RelativeOrAbsolute)
            };
        }
    }
}
