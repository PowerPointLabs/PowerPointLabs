using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorToMinimumSaturationHex : IValueConverter
    {

        public static SelectedColorToMinimumSaturationHex Instance = new SelectedColorToMinimumSaturationHex();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            HSLColor minSaturationHslColor = new HSLColor();
            minSaturationHslColor.Hue = selectedColor.Hue;
            minSaturationHslColor.Saturation = 0;
            minSaturationHslColor.Luminosity = selectedColor.Luminosity;
            Color minSaturationColor = minSaturationHslColor;
            return "#" + minSaturationColor.R.ToString("X2") + minSaturationColor.G.ToString("X2") + minSaturationColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
