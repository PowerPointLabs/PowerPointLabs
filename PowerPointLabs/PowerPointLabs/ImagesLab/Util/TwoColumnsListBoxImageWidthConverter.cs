using System;
using System.Globalization;
using System.Windows.Data;

namespace PowerPointLabs.ImagesLab.Util
{
    public class TwoColumnsListBoxImageWidthConverter : IValueConverter
    {
        private const double ImageMargin = 10;
        private const double ScrollBarMargin = 6;

        public object Convert(object value, Type targetType,
            object parameter, CultureInfo culture)
        {
            var originalValue = (double) value;
            return originalValue / 2 - ImageMargin - ScrollBarMargin;
        }

        public object ConvertBack(object value, Type targetType,
            object parameter, CultureInfo culture)
        {
            var valueAftConverted = (double) value;
            return (valueAftConverted + ImageMargin + ScrollBarMargin) * 2;
        }
    }
}
