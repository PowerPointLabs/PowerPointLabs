using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PowerPointLabs.Utils;

namespace PowerPointLabs.Extensions
{
    public static class Extension
    {

        public static void UpdateColors(this Control element, object sender, ColorTheme e)
        {
            // SolidColorBrush needs to be created on the same thread which the Control is created on.
            if (!element.Dispatcher.CheckAccess())
            {
                element.Dispatcher.Invoke(() => element.UpdateColors(sender, e));
                return;
            }
            element.Background = new SolidColorBrush(e.background);
            element.Foreground = new SolidColorBrush(e.foreground);

            foreach (Button button in element.GetElementType<Button>())
            {
                button.Background = new SolidColorBrush(e.background);
                button.Foreground = new SolidColorBrush(e.foreground);
            }
            foreach (CheckBox checkbox in element.GetElementType<CheckBox>())
            {
                checkbox.Foreground = new SolidColorBrush(e.foreground);
            }
            foreach (RadioButton radioButton in element.GetElementType<RadioButton>())
            {
                radioButton.Foreground = new SolidColorBrush(e.foreground);
            }
            foreach (TextBlock textBlock in element.GetElementType<TextBlock>())
            {
                textBlock.Foreground = new SolidColorBrush(e.foreground);
            }
            foreach (ListBox listView in element.GetElementType<ListBox>())
            {
                listView.Background = new SolidColorBrush(e.background);
                listView.Foreground = new SolidColorBrush(e.foreground);
            }
            foreach (Label label in element.GetElementType<Label>())
            {
                label.Foreground = new SolidColorBrush(e.foreground);
            }
            // ListView support required
        }

        /// <summary>
        /// Traverses the children of the parent and lazily returns children of type T.
        /// </summary>
        /// <typeparam name="T">Element type</typeparam>
        /// <param name="parent">Parent element of elements of interest</param>
        /// <returns></returns>
        public static IEnumerable<T> GetElementType<T>(this DependencyObject parent) where T : DependencyObject
        {
            foreach (T child in GetLogicalChildCollection<T>(parent))
            {
                yield return child;
            }
        }

        private static IEnumerable<T> GetLogicalChildCollection<T>(DependencyObject parent) where T : DependencyObject
        {
            foreach (object child in LogicalTreeHelper.GetChildren(parent))
            {
                if (child is DependencyObject)
                {
                    DependencyObject depChild = child as DependencyObject;
                    if (child is T)
                    {
                        yield return child as T;
                    }
                    foreach (T childOfChild in GetLogicalChildCollection<T>(depChild))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }
    }
}
