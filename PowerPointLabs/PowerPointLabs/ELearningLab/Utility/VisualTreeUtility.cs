using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.Utility
{
    public class VisualTreeUtility
    {

        public static List<Control> GetAllChildren(DependencyObject parent)
        {
            var list = new List<Control>();
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is Control)
                {
                    list.Add(child as Control);
                }
                list.AddRange(GetAllChildren(child));
            }
            return list;
        }

        public static FrameworkElement FindByName(string name, FrameworkElement root)
        {
            Stack<FrameworkElement> tree = new Stack<FrameworkElement>();
            tree.Push(root);

            while (tree.Count > 0)
            {
                FrameworkElement current = tree.Pop();
                if (current.Name == name)
                {
                    return current;
                }
                int count = VisualTreeHelper.GetChildrenCount(current);
                for (int i = 0; i < count; ++i)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(current, i);
                    if (child is FrameworkElement)
                    {
                        tree.Push((FrameworkElement)child);
                    }
                }
            }

            return null;
        }
    }
}
