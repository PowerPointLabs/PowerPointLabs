using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media.Imaging;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class AnimationTypeToImageSourceConverter : MarkupExtension, IValueConverter
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch ((AnimationType)value)
            {
                case AnimationType.Emphasis:
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.AnimationEmphasis.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
                case AnimationType.Entrance:
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.AnimationEntrance.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
                case AnimationType.Exit:
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.AnimationExit.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
                case AnimationType.MotionPath:
                default:
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.AnimationMotionPath.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            }
        }

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
