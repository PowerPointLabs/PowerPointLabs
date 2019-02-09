using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMaximumBrightnessHex : IValueConverter
    {
        public static SelectedColorToMaximumBrightnessHex Instance = new SelectedColorToMaximumBrightnessHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            HSLColor maxBrightnessHslColor = new HSLColor();
            maxBrightnessHslColor.Hue = selectedColor.Hue;
            maxBrightnessHslColor.Saturation = selectedColor.Saturation;
            maxBrightnessHslColor.Luminosity = 240;
            Color maxBrightnessColor = maxBrightnessHslColor;
            return "#" + maxBrightnessColor.R.ToString("X2") + maxBrightnessColor.G.ToString("X2") + maxBrightnessColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
