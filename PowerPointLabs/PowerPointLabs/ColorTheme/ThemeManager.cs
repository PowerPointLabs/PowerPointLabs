using System;
using System.Collections.Generic;
using System.Windows;
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
            // The ThemeWatcher will have no default keys.
            themeWatcher = new RegistryWatcher<int>(ThemeRegistryPath, ThemeRegistryKey, new List<int>());
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
                case ColorTheme.BLACK:
                case ColorTheme.COLORFUL:
                case ColorTheme.DARK_GRAY:
                case ColorTheme.DARK_GRAY_ALT:
                case ColorTheme.LIGHT_GRAY:
                case ColorTheme.WHITE:
                    _colorTheme.ThemeId = newValue;
                    break;
                default:
                    Logger.Log("Unknown UI Theme!");
                    _colorTheme.ThemeId = ColorTheme.COLORFUL;
                    break;
            }
        }
    }
}
