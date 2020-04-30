using PowerPointLabs.ActionFramework.Common.Log;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerPointLabs.ColorThemes
{
    /// <summary>
    /// The ThemeResourceDictionary class is a Resource Dictionary representing a Color Theme, to be added to the
    /// resources of a <see cref="FrameworkElement"/>.
    /// </summary>
    /// <remarks>
    /// The purpose of this class is to be able to uniquely identify the resource dictionary in a
    /// <see cref="FrameworkElement"/>'s resources that represents a Color Theme, so that it can be
    /// replaced when the Application's Color Theme changes.
    /// </remarks>
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
                    Logger.Log("Unknown UI Theme!");
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
