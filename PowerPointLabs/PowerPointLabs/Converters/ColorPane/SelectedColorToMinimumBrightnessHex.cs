using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMinimumBrightnessHex : IValueConverter
    {

        public static SelectedColorToMinimumBrightnessHex Instance = new SelectedColorToMinimumBrightnessHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            HSLColor minBrightnessHslColor = new HSLColor();
            minBrightnessHslColor.Hue = selectedColor.Hue;
            minBrightnessHslColor.Saturation = selectedColor.Saturation;
            minBrightnessHslColor.Luminosity = 0;
            Color minBrightnessColor = minBrightnessHslColor;
            return "#" + minBrightnessColor.R.ToString("X2") + minBrightnessColor.G.ToString("X2") + minBrightnessColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
