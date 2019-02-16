using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.Converters
{
#pragma warning disable 0618
    public class BlockToIndexConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ListViewItem item = (ListViewItem)value;
            ListView listView = ItemsControl.ItemsControlFromItemContainer(item) as ListView;
            int index = listView.ItemContainerGenerator.IndexFromContainer(item);
            ObservableCollection<ClickItem> items = listView.ItemsSource as ObservableCollection<ClickItem>;
            ClickItem clickItem = items.ElementAt(index);
            // TODO: Antipattern here. Try to think of a better way to get FirstClickNumber of slide
            CustomTaskPane eLearningTaskpane = Globals.ThisAddIn.GetActivePane(typeof(ELearningLabTaskpane));
            if (eLearningTaskpane == null)
            {
                return null;
            }
            ELearningLabTaskpane taskpane = eLearningTaskpane.Control as ELearningLabTaskpane;

            if (index == 0)
            {
                clickItem.ClickNo = taskpane.eLearningLabMainPanel1.FirstClickNumber;
            }
            else if (clickItem is SelfExplanationClickItem && (items.ElementAt(index - 1) is CustomClickItem)
                && (clickItem as SelfExplanationClickItem).TriggerIndex != (int)TriggerType.OnClick)
            {
                clickItem.ClickNo = items.ElementAt(index - 1).ClickNo;
            }
            else
            {
                clickItem.ClickNo = items.ElementAt(index - 1).ClickNo + 1;
            }
            return "Click " + clickItem.ClickNo;
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
