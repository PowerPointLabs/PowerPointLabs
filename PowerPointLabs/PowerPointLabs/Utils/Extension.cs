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
