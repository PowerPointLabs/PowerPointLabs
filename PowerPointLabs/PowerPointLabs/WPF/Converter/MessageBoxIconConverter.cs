using System;
using System.Windows;
using System.Windows.Data;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.WPF.Converter
{
    class MessageBoxIconConverter : BaseConverter, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
            {
                return DependencyProperty.UnsetValue;
            }

            return GetResourcePath((MessageBoxIcon)value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // not implemented
            return value;
        }

        public static string GetResourcePath(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Asterisk:
                    return "..\\Resources\\About.png";
                case MessageBoxIcon.Error:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Exclamation:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Question:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.None:
                default:
                    return "";
            }
        }
    }
}
