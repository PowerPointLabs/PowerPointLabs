using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

namespace PowerPointLabs.ColorThemes.Extensions
{
    public static class ThemeExtensions
    {

        public static bool? ShowThematicDialog(this Window w)
        {
            ThemeManager.Instance.ColorThemeChanged += w.ApplyTheme;
            bool? result = w.ShowDialog();
            ThemeManager.Instance.ColorThemeChanged -= w.ApplyTheme;
            return result;
        }

        public static void ApplyTheme(this DependencyObject element, object sender, ColorTheme theme)
        {
            if (!element.Dispatcher.CheckAccess())
            {
                element.Dispatcher.Invoke(() => element.ApplyTheme(sender, theme));
                return;
            }
            if (element.IsUpdated(theme)) { return; }
            switch (element)
            {
                case TextBlock t:
                    t.Foreground = new SolidColorBrush(theme.foreground);
                    break;
                case Label l:
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    break;
                case ListBox l:
                    l.Background = new SolidColorBrush(theme.background);
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    l.ResubscribeColorChangedHandler(sender, theme);
                    break;
                case Control c:
                    c.Background = new SolidColorBrush(theme.background);
                    c.Foreground = new SolidColorBrush(theme.foreground);
                    c.UpdateColorsControl(sender, theme);
                    break;
                case Panel p:
                    p.Background = new SolidColorBrush(theme.boxBackground);
                    p.UpdateColors(sender, theme);
                    break;
                case Border b:
                    b.Background = new SolidColorBrush(theme.boxBackground);
                    break;
                case Path p:
                    p.Stroke = new SolidColorBrush(theme.foreground);
                    break;
                default:
                    break;
            }
        }

        public static bool IsUpdated(this DependencyObject element, ColorTheme theme)
        {
            switch (element)
            {
                case TextBlock t:
                    return t.Foreground.IsBrushColor(theme.foreground);
                case Label l:
                    return l.Foreground.IsBrushColor(theme.foreground);
                case ListBox l:
                    return
                        l.Background.IsBrushColor(theme.background) &&
                        l.Foreground.IsBrushColor(theme.foreground);
                case Control c:
                    return
                        c.Background.IsBrushColor(theme.background) &&
                        c.Foreground.IsBrushColor(theme.foreground);
                case Panel p:
                    return p.Background.IsBrushColor(theme.boxBackground);
                case Border b:
                    return b.Background.IsBrushColor(theme.boxBackground);
                case Path p:
                    return p.Stroke.IsBrushColor(theme.foreground);
                default:
                    return true;
            }
        }

        private static bool IsBrushColor(this Brush b, Color color)
        {
            return b is SolidColorBrush && ((SolidColorBrush)b).Color == color;
        }

        // Uses VisualChildren, as some elements are ommitted if logical chlidren are used
        private static void UpdateColorsControl(this Control element, object sender, ColorTheme theme)
        {
            // For optimization reasons there is no need for the following line of code
            //element.ResubscribeColorChangedHandlerControl(sender, theme);
            foreach (Visual visual in GetVisualChildCollection<Visual>(element))
            {
                visual.ApplyTheme(sender, theme);
            }
        }

        // Uses LogicalChildren as it is much cheaper
        private static void UpdateColors(this DependencyObject element, object sender, ColorTheme theme)
        {
            foreach (DependencyObject dependencyObject in GetLogicalChildCollection<DependencyObject>(element))
            {
                dependencyObject.ApplyTheme(sender, theme);
            }
        }

        // Exploits the CommandBindings on Control to store actions to unsubscribe events
        private static void ResubscribeColorChangedHandler(this ListBox listBox, object sender, ColorTheme theme)
        {
            EventHandler StatusChangedHandler = new EventHandler((_o, _e) =>
            {
                if (listBox.ItemContainerGenerator.Status == GeneratorStatus.ContainersGenerated)
                {
                    listBox.UpdateColorsForChildren(sender, theme);
                }
            });
            EventHandler<DataTransferEventArgs> TargetUpdatedHandler = new EventHandler<DataTransferEventArgs>((_o, _e) =>
            {
                listBox.UpdateColorsForChildren(sender, theme);
            });
            ActionCommand command = new ActionCommand(() =>
            {
                listBox.ItemContainerGenerator.StatusChanged -= StatusChangedHandler;
                //listBox.TargetUpdated -= TargetUpdatedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in listBox.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            listBox.CommandBindings.Clear();
            listBox.UpdateColorsForChildren(sender, theme);
            listBox.ItemContainerGenerator.StatusChanged += StatusChangedHandler;
            //listBox.TargetUpdated += TargetUpdatedHandler;
            listBox.CommandBindings.Add(commandBinding);
        }

        private static void ResubscribeColorChangedHandlerControl(this Control control, object sender, ColorTheme theme)
        {
            EventHandler<DataTransferEventArgs> TargetUpdatedHandler = new EventHandler<DataTransferEventArgs>((_o, _e) =>
            {
                control.UpdateColorsControl(sender, theme);
            });
            ActionCommand command = new ActionCommand(() =>
            {
                control.TargetUpdated -= TargetUpdatedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in control.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            control.CommandBindings.Clear();
            control.TargetUpdated += TargetUpdatedHandler;
            control.CommandBindings.Add(commandBinding);
        }

        private static void UpdateColorsForChildren(this ListBox listBox, object sender, ColorTheme theme)
        {
            listBox.UpdateColors(sender, theme);
            for (int i = 0; i < listBox.Items.Count; i++)
            {
                Visual visual = listBox.ItemContainerGenerator.ContainerFromIndex(i) as Visual;
                if (visual == null) { break; }
                foreach (Visual element in GetVisualChildCollection<Visual>(visual))
                {
                    element.UpdateColors(sender, theme);
                }
            }
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

        private static IEnumerable<T> GetVisualChildCollection<T>(Visual parent) where T : Visual
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                Visual depChild = VisualTreeHelper.GetChild(parent, i) as Visual;
                if (depChild == null) continue;
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
