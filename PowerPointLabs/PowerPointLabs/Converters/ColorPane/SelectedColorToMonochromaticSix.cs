using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    class SelectedColorToMonochromaticSix : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var selectedColor = (HSLColor)value;
            Color convertedColor = new HSLColor(selectedColor.Hue, selectedColor.Saturation, 0.30f * 240);
            return Color.FromArgb(255,
                convertedColor.R,
                convertedColor.G,
                convertedColor.B);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
