using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using Color = System.Windows.Media.Color;

namespace PowerPointLabs.Converters.DrawingsLab
{
    public class ColorToBackgroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            byte r, g, b;
            Utils.Graphics.UnpackRgbInt((int)value, out r, out g, out b);
            return new SolidColorBrush(Color.FromRgb(r, g, b));

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var brush = value as SolidColorBrush;
            var color = brush.Color;
            return Utils.Graphics.PackRgbInt(color.R, color.G, color.B);
        }
    }
}
