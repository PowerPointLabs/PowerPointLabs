using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMaximumSaturationHex : IValueConverter
    {

        public static SelectedColorToMaximumSaturationHex Instance = new SelectedColorToMaximumSaturationHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            HSLColor maxSaturationHslColor = new HSLColor();
            maxSaturationHslColor.Hue = selectedColor.Hue;
            maxSaturationHslColor.Saturation = 240;
            maxSaturationHslColor.Luminosity = selectedColor.Luminosity;
            Color maxSaturationColor = maxSaturationHslColor;
            return "#" + maxSaturationColor.R.ToString("X2") + maxSaturationColor.G.ToString("X2") + maxSaturationColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
