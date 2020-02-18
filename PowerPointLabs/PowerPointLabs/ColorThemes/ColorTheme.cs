using System.Windows.Media;

namespace PowerPointLabs.ColorThemes
{
    public struct ColorTheme
    {
        public const int COLORFUL = 0;
        public const int LIGHT_GRAY = 1;
        public const int DARK_GRAY_ALT = 2;
        public const int DARK_GRAY = 3;
        public const int BLACK = 4;
        public const int WHITE = 5;

        public Color title;
        public Color background;
        public Color foreground;
        public Color boxBackground;
        public Color headingBackground;
        public Color headingForeground;

        /// <summary>
        /// The theme for Buttons used in this ColorTheme.
        /// </summary>
        public ButtonTheme ButtonTheme { get; set; }
    }
}
