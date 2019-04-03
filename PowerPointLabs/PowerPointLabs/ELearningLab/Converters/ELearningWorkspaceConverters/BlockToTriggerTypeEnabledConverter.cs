using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class BlockToTriggerTypeEnabledConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ListViewItem item = (ListViewItem)value;
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView;
            ObservableCollection<ClickItem> items = listView.ItemsSource as ObservableCollection<ClickItem>;
            int index = listView.ItemContainerGenerator.IndexFromContainer(item);
            ClickItem clickItem = items.ElementAt(index);
            if (clickItem is ExplanationItem && index > 0 && (items.ElementAt(index - 1) is CustomItem))
            {
                ((ExplanationItem)clickItem).IsTriggerTypeComboBoxEnabled = true;
                return true;
            }
            else
            {
                ((ExplanationItem)clickItem).IsTriggerTypeComboBoxEnabled = false;
                return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
