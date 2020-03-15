using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ColorThemes.Extensions
{
    public static class ThemeExtensions
    {
        public static readonly string PathToResourceDictionaries = "PowerPointLabs;component/Resources/ResourceDictionaries/";

        /// <summary>
        /// Shows a thematic dialog and waits for the window to close.
        /// </summary>
        /// <param name="window">Window to display</param>
        /// <param name="wait">Whether to wait for dialog to close</param>
        /// <returns></returns>
        public static bool? ShowThematicDialog(this Window window, bool wait = true)
        {
            window.Loaded -= window.RefreshVisual;
            window.Loaded += window.RefreshVisual;
            if (wait)
            {
                ThemeManager.Instance.ColorThemeChanged += window.ApplyTheme;
                bool? result = window.ShowDialog();
                ThemeManager.Instance.ColorThemeChanged -= window.ApplyTheme;
                return result;
            }
            else
            {
                ThemeManager.Instance.ColorThemeChanged += window.ApplyTheme;
                ThemeManager.Instance.ColorThemeChanged -= window.ApplyTheme;
                window.Show();
                return null;
            }
        }

        /// <summary>
        /// Invalidates the visual of the FrameworkElement.
        /// </summary>
        public static void RefreshVisual(this FrameworkElement element, object sender, RoutedEventArgs e)
        {
            element.InvalidateVisual();
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
            RemoveConflictingTheme(element);

            if (element is TextBox)
            {
                TextBox t = element as TextBox;
                t.Foreground = new SolidColorBrush(theme.foreground);
                t.Background = Brushes.Transparent;
                t.CaretBrush = new SolidColorBrush(theme.foreground);
            }
            else if (element is TextBlock)
            {
                TextBlock t = element as TextBlock;
                t.Foreground = new SolidColorBrush(theme.foreground);
                t.Background = Brushes.Transparent;
            }
            else if (element is CheckBox)
            {
                // The background of the CheckBox instance is actually the background of the box that
                // contains the checkmark (as opposed to the background of the label next to the box).
                // There is no simple way to change the color of the check mark itself, so to ensure
                // the check mark is stil visible, the background of the CheckBox shall remain unchanged.
                //
                // To change the check mark color, we will have to modify the ControlTemplate, which
                // is a pain. An example is given in the following link:
                // https://www.eidias.com/blog/2012/6/1/custom-wpf-check-box-with-inner-shadow-effect
                //
                // (I say that it's a pain, but it's what I did for Buttons)
                CheckBox t = element as CheckBox;
                t.Foreground = new SolidColorBrush(theme.foreground);
            }
            else if (element is Button)
            {
                Button t = element as Button;
                var style = CreateButtonStyleFromTheme(theme, t.Style);
                t.Style = style;
            }
            else if (element is Label)
            {
                Label l = element as Label;
                l.Foreground = new SolidColorBrush(theme.foreground);
                l.Background = Brushes.Transparent;
                l.BorderBrush = Brushes.Transparent;
            }
            else if (element is ListBox)
            {
                ListBox l = element as ListBox;
                l.Background = new SolidColorBrush(theme.background);
                l.Foreground = new SolidColorBrush(theme.foreground);
                l.ResubscribeColorChangedHandler(sender, theme);
            }
            else if (element is Frame)
            {
                Frame f = element as Frame;
                f.Background = new SolidColorBrush(theme.background);
                f.Foreground = new SolidColorBrush(theme.foreground);
                f.ResubscribeColorChangedHandler(sender, theme);
            }
            else if (element is Window)
            {
                Window w = element as Window;
                w.Background = new SolidColorBrush(theme.background);
                w.Foreground = new SolidColorBrush(theme.foreground);
                w.UpdateColors(sender, theme);
            }
            else if (element is Panel)
            {
                Panel p = element as Panel;
                p.Background = new SolidColorBrush(theme.background);
                p.UpdateColors(sender, theme); // the window is being update but doesn't show correctly
            }
            else if (element is Page)
            {
                Page p = element as Page;
                p.Foreground = new SolidColorBrush(theme.foreground);
                p.Background = new SolidColorBrush(theme.background);
                p.UpdateColorsVisual(sender, theme);
            }
            else if (element is Control)
            {
                Control c = element as Control;
                c.Background = new SolidColorBrush(theme.background);
                c.Foreground = new SolidColorBrush(theme.foreground);
                c.BorderBrush = new SolidColorBrush(theme.foreground);
                c.UpdateColorsVisual(sender, theme);
            }
            else if (element is Border)
            {
                // There is a problem with this approach, which is that every "Border",
                // regardless of what kind of Border it is, will have its background set.
                // This means you can't have a Border that has a custom Background that
                // differs from the theme. This problem appears for the other types as well.
                //
                // A possible solution would be to have all of these styles be placed in a
                // ResourceDictionary where the colors will be obtained via DynamicResource.
                // This way, all the controls will use these custom styles by default, but
                // can still be overridden when needed.
                //
                // There is a Border in SyncFormatListItem.xaml which has a White background
                // regardless of the theme. Hence, this line will be commented out for now.
                //
                // (element as Border).Background = new SolidColorBrush(theme.background);
            }
            else if (element is Path)
            {
                (element as Path).Stroke = new SolidColorBrush(theme.foreground);
            }
        }

        public static Control GetIWpfControl(this object control)
        {
            return (control as IWpfContainer)?.WpfControl;
        }

        /// <summary>
        /// Checks if an element has been updated with the current theme.
        /// </summary>
        /// <param name="element">WPF element to be updated</param>
        /// <param name="theme">Color theme</param>
        /// <returns></returns>
        public static bool IsUpdated(this DependencyObject element, ColorTheme theme)
        {
            if (element is TextBlock)
            {
                return (element as TextBlock).Foreground.IsBrushColor(theme.foreground);
            }
            else if (element is Label)
            {
                return (element as Label).Foreground.IsBrushColor(theme.foreground);
            }
            else if (element is Button)
            {
                Button b = element as Button;
                bool isMouseOver = b.IsMouseOver;
                bool isPressed = b.IsPressed;
                bool isDisabled = !b.IsEnabled;

                if (isDisabled)
                {
                    return b.Background.IsBrushColor(theme.ButtonTheme.DisabledBackground);
                }
                else if (isPressed)
                {
                    return b.Background.IsBrushColor(theme.ButtonTheme.PressedBackground);
                }
                else if (isMouseOver)
                {
                    return b.Background.IsBrushColor(theme.ButtonTheme.MouseOverBackground);
                }
                else
                {
                    return b.Background.IsBrushColor(theme.ButtonTheme.NormalBackground);
                }
            }
            else if (element is ListBox)
            {
                ListBox l = element as ListBox;
                return l.Background.IsBrushColor(theme.background) &&
                    l.Foreground.IsBrushColor(theme.foreground);
            }
            else if (element is Frame)
            {
                Frame f = element as Frame;
                return f.Background.IsBrushColor(theme.background) &&
                    f.Foreground.IsBrushColor(theme.foreground);
            }
            else if (element is Window)
            {
                Window w = element as Window;
                return w.Background.IsBrushColor(theme.background) &&
                    w.Foreground.IsBrushColor(theme.foreground);
            }
            else if (element is Panel)
            {
                return (element as Panel).Background.IsBrushColor(theme.boxBackground);
            }
            else if (element is Page)
            {
                Page p = element as Page;
                return p.Foreground.IsBrushColor(theme.foreground) &&
                    p.Background.IsBrushColor(theme.background);
            }
            else if (element is Control)
            {
                Control c = element as Control;
                return c.Background.IsBrushColor(theme.background) &&
                    c.Foreground.IsBrushColor(theme.foreground) &&
                    c.BorderBrush.IsBrushColor(theme.foreground);
            }
            else if (element is Border)
            {
                return (element as Border).Background.IsBrushColor(theme.boxBackground);
            }
            else if (element is Path)
            {
                return (element as Path).Stroke.IsBrushColor(theme.foreground);
            }
            return true;
        }

        // hotfix for combobox for AudioSettingsDialogWindow
        private static void RemoveConflictingTheme(DependencyObject element)
        {
            if (element is AudioSettingsDialogWindow)
            {
                AudioSettingsDialogWindow window = (AudioSettingsDialogWindow)element;
                Page p = window.MainPage;
                ResourceDictionary r = p.Resources.MergedDictionaries.FirstOrDefault(
                    (d) => d.Source.AbsoluteUri == "pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml");
                p.Resources.MergedDictionaries.Remove(r);
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
            EventHandler statusChangedHandler = new EventHandler((_o, _e) =>
            {
                if (frame.Content != null)
                {
                    frame.UpdateColorsChildren(sender, theme);
                }
            });
            ActionCommand command = new ActionCommand(() =>
            {
                frame.ContentRendered -= statusChangedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in frame.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            frame.CommandBindings.Clear();
            frame.ContentRendered += statusChangedHandler;
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
            EventHandler statusChangedHandler = new EventHandler((_o, _e) =>
            {
                if (listBox.ItemContainerGenerator.Status == GeneratorStatus.ContainersGenerated)
                {
                    listBox.UpdateColorsChildren(sender, theme);
                }
            });
            ActionCommand command = new ActionCommand(() =>
            {
                listBox.ItemContainerGenerator.StatusChanged -= statusChangedHandler;
            });
            CommandBinding commandBinding = new CommandBinding() { Command = command };
            foreach (CommandBinding binding in listBox.CommandBindings)
            {
                binding.Command.Execute(null);
            }
            listBox.CommandBindings.Clear();
            listBox.UpdateColorsChildren(sender, theme);
            listBox.ItemContainerGenerator.StatusChanged += statusChangedHandler;
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
                if (depChild == null)
                {
                    continue;
                }
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

        /// <summary>
        /// Retrieves the resource given by the specified key in the Resource Dictionary with the 
        /// specified name. The Resource Dictionary must be in the Resources/ResourceDictionaries
        /// directory.
        /// </summary>
        /// <param name="resourceDictionaryName">The name of the resource dictionary.</param>
        /// <param name="key">The key of the resource to retrieve.</param>
        /// <returns>The resouce to retrieve</returns>
        private static object GetResourceFromDictionary(string resourceDictionaryName, object key)
        {
            // TODO: Find out what happens when the key does not exist.
            var resourceDictionary = new ResourceDictionary
            {
                Source = new Uri(PathToResourceDictionaries + resourceDictionaryName + ".xaml", UriKind.RelativeOrAbsolute)
            };

            return resourceDictionary[key];
        }

        /// <summary>
        /// Construct a new Button Style that applies the given ColorTheme.
        /// </summary>
        /// <remarks>
        /// The basedOnStyle argument is used if the Style's target type is Button or a superclass 
        /// of Button. Otherwise, it will not be used.
        /// </remarks>
        /// <param name="theme">The ColorTheme to apply</param>
        /// <param name="basedOnStyle">The Button Style to base off of the returned Style</param>
        /// <returns>A Button Style that applies the given ColorTheme.</returns>
        private static Style CreateButtonStyleFromTheme(ColorTheme theme, Style basedOnStyle = null)
        {
            // To apply the behaviour of changing the background on mouse over (and other 
            // conditions), we need to overwrite the Control Template. It's troublesome to create 
            // the entire Control Template programmatically, so we'll create as much as we can in
            // XAML.
            var controlTemplate = GetResourceFromDictionary("GeneralResourceDictionary", "DefaultButtonControlTemplate") as ControlTemplate;

            // Set the behaviour for when the mouse is over the Button.
            var mouseOverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true,
            };
            mouseOverTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(theme.ButtonTheme.MouseOverBackground)));
            mouseOverTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(theme.ButtonTheme.MouseOverBorderColor)));

            // Set the behaviour for when the Button is pressed.
            var pressedTrigger = new Trigger
            {
                Property = ButtonBase.IsPressedProperty,
                Value = true
            };
            pressedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(theme.ButtonTheme.PressedBackground)));
            pressedTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(theme.ButtonTheme.PressedBorderColor)));

            // Set the behaviour for when the Button is disabled.
            var disabledTrigger = new Trigger
            {
                Property = UIElement.IsEnabledProperty,
                Value = false
            };
            disabledTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(theme.ButtonTheme.DisabledBackground)));
            disabledTrigger.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(theme.ButtonTheme.DisabledForeground)));
            disabledTrigger.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(theme.ButtonTheme.DisabledBorderColor)));

            // Apply the triggers to the Control Template.
            controlTemplate.Triggers.Add(mouseOverTrigger);
            controlTemplate.Triggers.Add(pressedTrigger);
            controlTemplate.Triggers.Add(disabledTrigger);

            // Create the output style. 
            var outputStyle = new Style(typeof(Button));
            // Base it off of the specified base Style, but only if compatible.
            if (basedOnStyle != null && basedOnStyle.TargetType.IsAssignableFrom(typeof(Button)))
            {
                outputStyle.BasedOn = basedOnStyle;
            };

            outputStyle.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(theme.ButtonTheme.NormalBackground)));
            outputStyle.Setters.Add(new Setter(Control.ForegroundProperty, new SolidColorBrush(theme.ButtonTheme.NormalForeground)));
            outputStyle.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(theme.ButtonTheme.NormalBorderColor)));
            outputStyle.Setters.Add(new Setter(Control.TemplateProperty, controlTemplate));

            return outputStyle;
        }
    }
}
