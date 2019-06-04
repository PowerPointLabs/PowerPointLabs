using System;
using System.Collections.Generic;
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
        /// <summary>
        /// Shows a thematic dialog and waits for the window to close.
        /// </summary>
        /// <param name="window">Window to display</param>
        /// <param name="wait">Whether to wait for dialog to close</param>
        /// <returns></returns>
        public static bool? ShowThematicDialog(this Window window, bool wait = true)
        {
            if (wait)
            {
                ThemeManager.Instance.ColorThemeChanged += window.ApplyTheme;
                bool? result = window.ShowDialog();
                ThemeManager.Instance.ColorThemeChanged -= window.ApplyTheme;
                return result;
            }
            else
            {
                window.ApplyTheme(null, ThemeManager.Instance.ColorTheme);
                window.Show();
                return null;
            }
        }

        /// <summary>
        /// Applies a theme to the element recursively.
        /// </summary>
        /// <param name="element">WPF element to apply theme to.</param>
        /// <param name="sender">Object that triggered the event</param>
        /// <param name="theme">Color theme to update with.</param>
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
                // textbox placeholder vanishes
                case TextBlock t:
                    t.Foreground = new SolidColorBrush(theme.foreground);
                    t.Background = Brushes.Transparent;
                    break;
                case Label l:
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    l.Background = Brushes.Transparent;
                    l.BorderBrush = Brushes.Transparent;
                    break;
                case ListBox l:
                    l.Background = new SolidColorBrush(theme.background);
                    l.Foreground = new SolidColorBrush(theme.foreground);
                    l.ResubscribeColorChangedHandler(sender, theme);
                    break;
                case Frame f:
                    f.Background = new SolidColorBrush(theme.background);
                    f.Foreground = new SolidColorBrush(theme.foreground);
                    f.ResubscribeColorChangedHandler(sender, theme);
                    break;
                case Window w:
                    w.Background = new SolidColorBrush(theme.background);
                    w.Foreground = new SolidColorBrush(theme.foreground);
                    w.UpdateColors(sender, theme);
                    break;
                case Panel p:
                    p.Background = new SolidColorBrush(theme.boxBackground);
                    p.UpdateColors(sender, theme); // the window is being update but doesn't show correctly
                    break;
                case Page p:
                    p.Foreground = new SolidColorBrush(theme.foreground);
                    p.Background = new SolidColorBrush(theme.background);
                    p.UpdateColorsVisual(sender, theme);
                    break;
                case Control c:
                    c.Background = new SolidColorBrush(theme.background);
                    c.Foreground = new SolidColorBrush(theme.foreground);
                    c.BorderBrush = new SolidColorBrush(theme.foreground);
                    c.UpdateColorsVisual(sender, theme);
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

        /// <summary>
        /// Checks if an element has been updated with the current theme.
        /// </summary>
        /// <param name="element">WPF element to be updated</param>
        /// <param name="theme">Color theme</param>
        /// <returns></returns>
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
                case Frame f:
                    return
                        f.Background.IsBrushColor(theme.background) &&
                        f.Foreground.IsBrushColor(theme.foreground);
                case Window w:
                    return
                        w.Background.IsBrushColor(theme.background) &&
                        w.Foreground.IsBrushColor(theme.foreground);
                case Panel p:
                    return p.Background.IsBrushColor(theme.boxBackground);
                case Page p:
                    return
                        p.Foreground.IsBrushColor(theme.foreground) &&
                        p.Background.IsBrushColor(theme.background);
                case Control c:
                    return
                        c.Background.IsBrushColor(theme.background) &&
                        c.Foreground.IsBrushColor(theme.foreground) &&
                        c.BorderBrush.IsBrushColor(theme.foreground);
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
        private static void UpdateColorsVisual(this FrameworkElement element, object sender, ColorTheme theme)
        {
            foreach (Visual visual in GetVisualChildCollection<Visual>(element))
            {
                visual.ApplyTheme(sender, theme);
            }
            element.ApplyTemplate();
        }

        // Uses LogicalChildren as it is much cheaper
        private static void UpdateColors(this DependencyObject element, object sender, ColorTheme theme)
        {
            foreach (DependencyObject dependencyObject in GetLogicalChildCollection<DependencyObject>(element))
            {
                dependencyObject.ApplyTheme(sender, theme);
            }
        }

        private static void ResubscribeColorChangedHandler(this Frame frame, object sender, ColorTheme theme)
        {
            EventHandler StatusChangedHandler = new EventHandler((_o, _e) =>
            {
                if (frame.Content != null)
                {
                    frame.UpdateColorsChildren(sender, theme);
                }
            });
            ActionCommand command = new ActionCommand(() =>
            {
                frame.ContentRendered -= StatusChangedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in frame.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            frame.CommandBindings.Clear();
            frame.ContentRendered += StatusChangedHandler;
            frame.CommandBindings.Add(commandBinding);
        }

        private static void UpdateColorsChildren(this Frame frame, object sender, ColorTheme theme)
        {
            Visual visual = frame.Content as Visual;
            if (visual == null) { return; }
            foreach (Visual element in GetVisualChildCollection<Visual>(visual))
            {
                element.UpdateColors(sender, theme);
            }
            visual.UpdateColors(sender, theme);
        }

        // Exploits the CommandBindings on Control to store actions to unsubscribe events
        private static void ResubscribeColorChangedHandler(this ListBox listBox, object sender, ColorTheme theme)
        {
            EventHandler StatusChangedHandler = new EventHandler((_o, _e) =>
            {
                if (listBox.ItemContainerGenerator.Status == GeneratorStatus.ContainersGenerated)
                {
                    listBox.UpdateColorsChildren(sender, theme);
                }
            });
            ActionCommand command = new ActionCommand(() =>
            {
                listBox.ItemContainerGenerator.StatusChanged -= StatusChangedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in listBox.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            listBox.CommandBindings.Clear();
            listBox.UpdateColorsChildren(sender, theme);
            listBox.ItemContainerGenerator.StatusChanged += StatusChangedHandler;
            listBox.CommandBindings.Add(commandBinding);
        }

        private static void UpdateColorsChildren(this ListBox listBox, object sender, ColorTheme theme)
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
                visual.UpdateColors(sender, theme);
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
