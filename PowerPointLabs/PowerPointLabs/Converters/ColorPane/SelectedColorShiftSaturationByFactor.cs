using System;
using System.Drawing;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

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
                return ColorHelper.ColorToHexString(color);
            }
            HSLColor selectedColor = (HSLColor)value;
            float shiftFactor = float.Parse((string)parameter);
            Color convertedColor = new HSLColor(selectedColor.Hue, shiftFactor * 240, selectedColor.Luminosity);
            return ColorHelper.ColorToHexString(convertedColor);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
