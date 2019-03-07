using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class ListViewClickNoToColorConverter : MarkupExtension, IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool isDummyItem = (bool)values[0];
            bool isTriggerTypeEnabled = (bool)values[1];
            int triggerType = (int)values[2];
            if (isDummyItem)
            {
                return new SolidColorBrush(Colors.Gray);
            }
            else if (isTriggerTypeEnabled && triggerType == (int)TriggerType.WithPrevious)
            {
                return new SolidColorBrush(Colors.Transparent);
            }
            else
            {
                return new SolidColorBrush(Colors.Black);
            }
        }


        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
