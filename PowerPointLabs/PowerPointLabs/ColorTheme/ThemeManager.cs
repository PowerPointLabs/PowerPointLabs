using System;
using System.Windows.Media;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ColorThemes
{
    /// <summary>
    /// A class that manages the changing of colors to match the UI theme.
    /// </summary>
    public class ThemeManager
    {
        public static ThemeManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ThemeManager();
                }
                return _instance;
            }
        }
        public static string ThemeRegistryPath
        {
            get
            {
                return String.Format(@"SOFTWARE\\Microsoft\\Office\\{0}\\Common", Globals.ThisAddIn.Application.Version);
            }
        }
        public readonly string ThemeRegistryKey = "UI Theme";
        public static void TearDown()
        {
            if (_instance == null)
            {
                return;
            }
            _instance.themeWatcher.Stop();
            _instance = null;
        }
        private static ThemeManager _instance;

        public event EventHandler<ColorTheme> ColorThemeChanged
        {
            add
            {
                value(this, _colorTheme);
                _ColorThemeChanged += value;
            }
            remove
            {
                _ColorThemeChanged -= value;
            }
        }
        public ColorTheme ColorTheme => _colorTheme;

        private RegistryWatcher<int> themeWatcher;
        private event EventHandler<ColorTheme> _ColorThemeChanged;
        private ColorTheme _colorTheme;

        private ThemeManager()
        {
            themeWatcher = new RegistryWatcher<int>(ThemeRegistryPath, ThemeRegistryKey);
            themeWatcher.ValueChanged += ThemeChangedHandler;
            themeWatcher.Fire();
            themeWatcher.Start();
        }

        private void ThemeChangedHandler(object sender, int newValue)
        {
            UpdateColorTheme(newValue);
            _ColorThemeChanged(this, _colorTheme);
        }

        private void UpdateColorTheme(int newValue)
        {
            switch (newValue)
            {
                case ColorTheme.COLORFUL:
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(230, 230, 230);
                    _colorTheme.foreground = Color.FromRgb(37, 37, 37);
                    _colorTheme.boxBackground = Color.FromRgb(255, 255, 255);
                    _colorTheme.headingBackground = Color.FromRgb(181, 71, 42);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    break;
                case ColorTheme.WHITE:
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(255, 255, 255);
                    _colorTheme.foreground = Color.FromRgb(37, 37, 37);
                    _colorTheme.boxBackground = Color.FromRgb(230, 230, 230);
                    _colorTheme.headingBackground = Color.FromRgb(181, 71, 42);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    break;
                case ColorTheme.DARK_GREY:
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(102, 102, 102);
                    _colorTheme.foreground = Color.FromRgb(238, 238, 238);
                    _colorTheme.boxBackground = Color.FromRgb(64, 64, 64);
                    _colorTheme.headingBackground = Color.FromRgb(208, 71, 38);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    break;
                case ColorTheme.BLACK:
                    _colorTheme.title = Color.FromRgb(239, 239, 239);
                    _colorTheme.background = Color.FromRgb(37, 37, 37);
                    _colorTheme.foreground = Color.FromRgb(238, 238, 238);
                    _colorTheme.boxBackground = Color.FromRgb(64, 64, 64);
                    _colorTheme.headingBackground = Color.FromRgb(208, 71, 38);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    break;
                default:
                    Logger.Log("Unknown UI Theme!");
                    break;
            }
        }
    }
}
