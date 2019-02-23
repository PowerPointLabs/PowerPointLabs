using System;
using System.Windows.Data;

using PowerPointLabs.ColorsLab;

namespace PowerPointLabs.Converters.ColorPane
{
    [ValueConversion(typeof(HSLColor), typeof(int))]
    class SelectedColorToSaturationValue : IValueConverter
    {

        public static SelectedColorToSaturationValue Instance = new SelectedColorToSaturationValue();

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                return 0;
            }

            HSLColor selectedColor = (HSLColor)value;
            return (int)(selectedColor.Saturation);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
