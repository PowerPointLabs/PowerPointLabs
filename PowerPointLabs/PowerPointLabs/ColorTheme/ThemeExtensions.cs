using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace PowerPointLabs.ColorThemes.Extensions
{
    public static class ThemeExtensions
    {

        public static bool? ShowThematicDialog(this Window w)
        {
            ThemeManager.Instance.ColorThemeChanged += w.UpdateColorsControl;
            bool? result = w.ShowDialog();
            ThemeManager.Instance.ColorThemeChanged -= w.UpdateColorsControl;
            return result;
        }

        public static void UpdateColorsControl(this Control element, object sender, ColorTheme theme)
        {
            if (!element.Dispatcher.CheckAccess())
            {
                element.Dispatcher.Invoke(() => element.UpdateColorsControl(sender, theme));
                return;
            }
            element.Background = new SolidColorBrush(theme.background);
            element.Foreground = new SolidColorBrush(theme.foreground);
            foreach (DependencyObject o in GetVisualChildCollection<DependencyObject>(element))
            {
                o.ApplyTheme(sender, theme);
            }
        }

        public static void ApplyTheme(this DependencyObject element, object sender, ColorTheme theme)
        {
            switch (element)
            {
                //case ComboBox c:
                //    break;
                //case ToggleButton b:
                //    b.Foreground = new SolidColorBrush(theme.foreground);
                //    break;
                case ListBox l:
                    l.Background = new SolidColorBrush(theme.background);
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    ResubscribeAlt(sender, theme, l);
                    break;
                case Panel p:
                    p.Background = new SolidColorBrush(theme.boxBackground);
                    p.UpdateColors(sender, theme);
                    break;
                case Border b:
                    b.Background = new SolidColorBrush(theme.boxBackground);
                    break;
                case TextBlock t:
                    t.Foreground = new SolidColorBrush(theme.foreground);
                    break;
                case Label l:
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    break;
                case ContentControl c:
                    c.UpdateColorsContentControl(sender, theme);
                    break;
                case Control c:
                    c.UpdateColorsControl(sender, theme);
                    break;
                default:
                    break;
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

        private static void UpdateColors(this DependencyObject element, object sender, ColorTheme theme)
        {
            if (!element.Dispatcher.CheckAccess())
            {
                element.Dispatcher.Invoke(() => element.UpdateColors(sender, theme));
                return;
            }
            foreach (DependencyObject o in GetLogicalChildCollection<DependencyObject>(element))
            {
                o.ApplyTheme(sender, theme);
            }
        }

        private static void UpdateColorsContentControl(this ContentControl element, object sender, ColorTheme theme)
        {
            if (!element.Dispatcher.CheckAccess())
            {
                element.Dispatcher.Invoke(() => element.UpdateColorsContentControl(sender, theme));
                return;
            }
            element.Background = new SolidColorBrush(theme.background);
            element.Foreground = new SolidColorBrush(theme.foreground);
            foreach (DependencyObject o in GetVisualChildCollection<DependencyObject>(element))
            {
                o.ApplyTheme(sender, theme);
            }
        }

        // Exploits the CommandBindings on Control to store actions to unsubscribe events
        private static void ResubscribeAlt(object sender, ColorTheme theme, ListBox listBox)
        {
            EventHandler h = new EventHandler((_o, _e) =>
            {
                if (listBox.ItemContainerGenerator.Status == GeneratorStatus.ContainersGenerated)
                {
                    listBox.UpdateColors(sender, theme);
                    for (int i = 0; i < listBox.Items.Count; i++)
                    {
                        DependencyObject o = listBox.ItemContainerGenerator.ContainerFromIndex(i);
                        if (o == null) { break; }
                        foreach (DependencyObject element in GetVisualChildCollection<DependencyObject>(o))
                        {
                            element.UpdateColors(sender, theme);
                        }
                    }
                }
            });
            ActionCommand command = new ActionCommand(() => listBox.ItemContainerGenerator.StatusChanged -= h);
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in listBox.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            listBox.CommandBindings.Clear();
            h(sender, null);
            listBox.ItemContainerGenerator.StatusChanged += h;
            listBox.CommandBindings.Add(commandBinding);
        }

        [Obsolete]
        private static void ResubscribeColorEvent(object sender, ColorTheme e, ListBox listBox)
        {
            NotifyCollectionChangedEventHandler updateListBox = (_sender, _e) =>
            {
                listBox.UpdateColors(sender, e);
            };
            INotifyCollectionChanged items = listBox.Items;
            ActionCommand command = new ActionCommand(() => items.CollectionChanged -= updateListBox);
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in listBox.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            listBox.CommandBindings.Clear();
            items.CollectionChanged += updateListBox;
            listBox.CommandBindings.Add(commandBinding);
        }

        private static IEnumerable<T> GetChildCollection<T>(DependencyObject parent) where T : DependencyObject
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
                    foreach (T childOfChild in GetChildCollection<T>(depChild))
                    {
                        yield return childOfChild;
                    }
                }
            }

            if (parent is Visual || parent is Visual3D)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
                {
                    DependencyObject depChild = VisualTreeHelper.GetChild(parent, i);
                    if (depChild is T)
                    {
                        yield return depChild as T;
                    }
                    foreach (T childOfChild in GetChildCollection<T>(depChild))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        private static IEnumerable<T> GetVisualChildCollection<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent is Visual || parent is Visual3D)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
                {
                    DependencyObject depChild = VisualTreeHelper.GetChild(parent, i);
                    if (depChild is T)
                    {
                        yield return depChild as T;
                    }
                    foreach (T childOfChild in GetVisualChildCollection<T>(depChild))
                    {
                        yield return childOfChild;
                    }
                }
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
