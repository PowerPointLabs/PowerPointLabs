using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerPointLabs.Resources.Themes
{
    /// <summary>
    /// The DesignTheme class is a ResourceDictionary for use in the Designer in Visual Studio.
    /// </summary>
    /// <remarks>
    /// This class will apply the specified theme in the Designer preview, allowing the programmer
    /// to design Windows or User Control .xaml files while being able to freely change the themes.
    /// This ResourceDictionary will have no effect during runtime.
    /// </remarks>
    public class DesignTheme : ResourceDictionary
    {
        public static readonly string PathToThemesFolder = "pack://application:,,,/PowerPointLabs;component/Resources/Themes/";

        private string theme;

        /// <summary>
        /// The theme to apply. Valid values are "Colorful", "Black", "White" and "Dark Gray".
        /// </summary>
        public string Theme
        {
            get => theme;
            set
            {
                theme = value;

                // Set the Source of the Resource Dictionary only when in Designer mode(and not runtime).
                if ((bool)DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(DependencyObject)).DefaultValue)
                {
                    var themeUri = PathToThemesFolder + theme + "Theme.xaml";
                    Source = new Uri(themeUri);
                }
            }
        }
    }
}
