﻿using System.Windows.Media;

namespace PowerPointLabs.ColorThemes
{
    public struct ColorTheme
    {
        public const int COLORFUL = 0;
        public const int DARK_GREY = 3;
        public const int BLACK = 4;
        public const int WHITE = 5;

        public Color title;
        public Color background;
        public Color foreground;
        public Color headingBackground;
        public Color headingForeground;
    }
}
