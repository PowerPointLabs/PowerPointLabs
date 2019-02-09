using System;
using System.Drawing;
using System.Windows.Data;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(string))]
    class SelectedColorShiftSaturationByFactor : IValueConverter
    {
        public static SelectedColorShiftSaturationByFactor Instance = new SelectedColorShiftSaturationByFactor();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                Color color = (HSLColor)value;
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
            HSLColor selectedColor = (HSLColor)value;
            float shiftFactor = float.Parse((string)parameter);
            Color convertedColor = new HSLColor(selectedColor.Hue, shiftFactor * 240, selectedColor.Luminosity);
            return "#" + convertedColor.R.ToString("X2") + convertedColor.G.ToString("X2") + convertedColor.B.ToString("X2");
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
