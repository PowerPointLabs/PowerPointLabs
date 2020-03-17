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
                if (!themeWatcher.IsDefaultKey)
                {
                    value(this, _colorTheme);
                }
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
            themeWatcher = new RegistryWatcher<int>(ThemeRegistryPath, ThemeRegistryKey, GetDefaultKeys());
            themeWatcher.ValueChanged += ThemeChangedHandler;
            themeWatcher.Fire();
            themeWatcher.Start();
        }

        private List<int> GetDefaultKeys()
        {
            if (!Globals.ThisAddIn.IsApplicationVersion2013())
            {
                return new List<int>() { ColorTheme.COLORFUL };
            }
            return new List<int>()
            {
                ColorTheme.WHITE,
                ColorTheme.LIGHT_GRAY,
                ColorTheme.DARK_GRAY
            };
        }

        private void ThemeChangedHandler(object sender, int newValue)
        {
            UpdateColorTheme(newValue);
            _ColorThemeChanged?.Invoke(this, _colorTheme);
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
                    _colorTheme.ButtonTheme = new ButtonTheme
                    {
                        // The following values are obtained by copying the appearance
                        // of the "Play All" button from the (standard) PowerPoint 
                        // Animation pane.
                        NormalBackground = Color.FromRgb(253, 253, 253),
                        NormalForeground = _colorTheme.foreground,
                        NormalBorderColor = Color.FromRgb(171, 171, 171),
                        MouseOverBackground = Color.FromRgb(252, 228, 220),
                        MouseOverBorderColor = Color.FromRgb(245, 186, 157),
                        PressedBackground = Color.FromRgb(245, 186, 157),
                        PressedBorderColor = Color.FromRgb(240, 98, 62),
                        DisabledBackground = Color.FromRgb(253, 253, 253),
                        DisabledForeground = Color.FromRgb(204, 177, 192),
                        DisabledBorderColor = Color.FromRgb(225, 225, 225)
                    };
                    break;
                case ColorTheme.WHITE:
                case ColorTheme.LIGHT_GRAY:
                case ColorTheme.DARK_GRAY_ALT:
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(255, 255, 255);
                    _colorTheme.foreground = Color.FromRgb(37, 37, 37);
                    _colorTheme.boxBackground = Color.FromRgb(230, 230, 230);
                    _colorTheme.headingBackground = Color.FromRgb(181, 71, 42);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    _colorTheme.ButtonTheme = new ButtonTheme
                    {
                        NormalBackground = Color.FromRgb(212, 212, 212),
                        NormalForeground = Color.FromRgb(38, 38, 38),
                        NormalBorderColor = Color.FromRgb(87, 87, 87),
                        MouseOverBackground = Color.FromRgb(249, 201, 185),
                        MouseOverBorderColor = Color.FromRgb(235, 117, 59),
                        PressedBackground = Color.FromRgb(235, 117, 59),
                        PressedBorderColor = Color.FromRgb(225, 0, 0),
                        DisabledBackground = Color.FromRgb(212, 212, 212),
                        DisabledForeground = Color.FromRgb(172, 152, 162),
                        DisabledBorderColor = Color.FromRgb(195, 195, 195)
                    };
                    break;
                case ColorTheme.DARK_GRAY:
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(102, 102, 102);
                    _colorTheme.foreground = Color.FromRgb(238, 238, 238);
                    _colorTheme.boxBackground = Color.FromRgb(64, 64, 64);
                    _colorTheme.headingBackground = Color.FromRgb(208, 71, 38);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    _colorTheme.ButtonTheme = new ButtonTheme
                    {
                        NormalBackground = Color.FromRgb(212, 212, 212),
                        NormalForeground = Color.FromRgb(38, 38, 38),
                        NormalBorderColor = Color.FromRgb(87, 87, 87),
                        MouseOverBackground = Color.FromRgb(249, 201, 185),
                        MouseOverBorderColor = Color.FromRgb(235, 117, 59),
                        PressedBackground = Color.FromRgb(235, 117, 59),
                        PressedBorderColor = Color.FromRgb(225, 0, 0),
                        DisabledBackground = Color.FromRgb(212, 212, 212),
                        DisabledForeground = Color.FromRgb(172, 152, 162),
                        DisabledBorderColor = Color.FromRgb(195, 195, 195)
                    };
                    break;
                case ColorTheme.BLACK:
                    _colorTheme.title = Color.FromRgb(239, 239, 239);
                    _colorTheme.background = Color.FromRgb(37, 37, 37);
                    _colorTheme.foreground = Color.FromRgb(238, 238, 238);
                    _colorTheme.boxBackground = Color.FromRgb(64, 64, 64);
                    _colorTheme.headingBackground = Color.FromRgb(208, 71, 38);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    _colorTheme.ButtonTheme = new ButtonTheme
                    {
                        NormalBackground = Color.FromRgb(68, 68, 68),
                        NormalForeground = _colorTheme.foreground,
                        NormalBorderColor = Color.FromRgb(106, 106, 106),
                        MouseOverBackground = Color.FromRgb(68, 68, 68),
                        MouseOverBorderColor = Color.FromRgb(150, 150, 150),
                        PressedBackground = Color.FromRgb(106, 106, 106),
                        PressedBorderColor = Color.FromRgb(150, 150, 150),
                        DisabledBackground = Color.FromRgb(37, 37, 37),
                        DisabledForeground = Color.FromRgb(73, 90, 82),
                        DisabledBorderColor = Color.FromRgb(68, 68, 68)
                    };
                    break;
                default:
                    Logger.Log("Unknown UI Theme!");
                    _colorTheme.title = Color.FromRgb(181, 71, 42);
                    _colorTheme.background = Color.FromRgb(230, 230, 230);
                    _colorTheme.foreground = Color.FromRgb(37, 37, 37);
                    _colorTheme.boxBackground = Color.FromRgb(255, 255, 255);
                    _colorTheme.headingBackground = Color.FromRgb(181, 71, 42);
                    _colorTheme.headingForeground = Color.FromRgb(238, 238, 238);
                    _colorTheme.ButtonTheme = new ButtonTheme
                    {
                        NormalBackground = Color.FromRgb(221, 221, 221),
                        NormalForeground = _colorTheme.foreground,
                        NormalBorderColor = Color.FromRgb(112, 112, 112),
                        MouseOverBackground = Color.FromRgb(190, 230, 253),
                        MouseOverBorderColor = Color.FromRgb(60, 127, 177),
                        PressedBackground = Color.FromRgb(196, 229, 246),
                        PressedBorderColor = Color.FromRgb(44, 98, 139),
                        DisabledBackground = Color.FromRgb(224, 224, 224),
                        DisabledForeground = Color.FromRgb(131, 131, 131),
                        DisabledBorderColor = Color.FromRgb(173, 178, 181)
                    };
                    break;
            }
        }
    }
}
