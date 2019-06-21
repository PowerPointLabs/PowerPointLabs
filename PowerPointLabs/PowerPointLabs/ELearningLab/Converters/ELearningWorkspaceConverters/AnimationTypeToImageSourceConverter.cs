using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media.Imaging;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.Utils;

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
                    return GraphicsUtil.CreateBitmapSource(Properties.Resources.AnimationEmphasis);
                case AnimationType.Entrance:
                    return GraphicsUtil.CreateBitmapSource(Properties.Resources.AnimationEntrance);
                case AnimationType.Exit:
                    return GraphicsUtil.CreateBitmapSource(Properties.Resources.AnimationExit);
                case AnimationType.MotionPath:
                default:
                    return GraphicsUtil.CreateBitmapSource(Properties.Resources.AnimationMotionPath);
            }
        }

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
