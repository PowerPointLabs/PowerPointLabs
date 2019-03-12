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
using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;

namespace PowerPointLabs.ELearningLab.Converters
{
#pragma warning disable 0618
    public class ListItemClickNumToColorConverter : MarkupExtension, IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool isDummyItem, isTriggerTypeEnabled;
            int triggerType, clickNum;
            try
            {
                isDummyItem = (bool)values[0];
                isTriggerTypeEnabled = (bool)values[1];
                triggerType = (int)values[2];
                clickNum = (int)values[3];
            }
            catch (Exception e)
            {
                Logger.Log("Invalid Cast Exception" + e.Message);
                return null;
            }
            CustomTaskPane eLearningTaskpane = Globals.ThisAddIn.GetActivePane(typeof(ELearningLabTaskpane));
            if (eLearningTaskpane == null)
            {
                return null;
            }
            ELearningLabTaskpane taskpane = eLearningTaskpane.Control as ELearningLabTaskpane;
            if (isDummyItem)
            {
                return new SolidColorBrush(Colors.Gray);
            }
            else if (isTriggerTypeEnabled && triggerType == (int)TriggerType.WithPrevious
                && !taskpane.ELearningLabMainPanel.IsFirstItemSelfExplanation && clickNum == 0)
            {
                return new SolidColorBrush(Colors.Transparent);
            }
            else if (isTriggerTypeEnabled && triggerType == (int)TriggerType.WithPrevious && clickNum > 0)
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
