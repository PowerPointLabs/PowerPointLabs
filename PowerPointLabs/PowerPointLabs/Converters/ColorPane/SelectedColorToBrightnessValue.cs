using System;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(int))]
    class SelectedColorToBrightnessValue : IValueConverter
    {

        public static SelectedColorToBrightnessValue Instance = new SelectedColorToBrightnessValue();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                return 0;
            }

            HSLColor selectedColor = (HSLColor)value;
            return (int)(selectedColor.Luminosity);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
