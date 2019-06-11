using System;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MahApps.Metro.Controls;
using TestInterface.Windows;

namespace PowerPointLabs.FunctionalTestInterface.Windows
{
    public class MarshalWindow : MarshalByRefObject, IMarshalWindow
    {
        private readonly Window window;

        public string Title => BlockUntilSTAThread(() => window.Title);

        private MarshalWindow(Window w)
        {
            window = w;
        }

        public static MarshalWindow CreateInstance(Window w)
        {
            if (w == null) { return null; }
            return new MarshalWindow(w);
        }

        public void RaiseEvent<T>(string name, RoutedEventArgs args)
        {
            BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                element.RaiseEvent(args);
            });
        }

        public bool Focus<T>(string name)
        {
            return BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                element.Focusable = true;
                return element.Focus();
            });
        }

        public void SelectAll<T>(string name)
        {
            BlockUntilSTAThread(() =>
            {
                TextBox element = GetElement<T>(name) as TextBox;
                element?.SelectAll();
            });
        }

        public void LeftClick<T>(string name)
        {
            BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                PresentationSource source = PresentationSource.FromVisual(element);

                element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = UIElement.MouseDownEvent
                });
                Thread.Sleep(30);
                element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = UIElement.MouseUpEvent
                });
            });
        }

        public Point GetListElementPosition<T>(string name, int index)
        {
            return BlockUntilSTAThread(() =>
            {
                Control element = GetElement<T>(name) as Control;
                if (element == null) { return new Point(0, 0); }
                int factor = 3 + 2 * index;
                return element.PointToScreen(new Point(element.ActualWidth / 2, element.ActualHeight * factor / 2));
            });
            //return BlockUntilSTAThread(() =>
            //{
            //    ItemsControl element = GetElement<T>(name) as ItemsControl;
            //    if (element == null) { return new Point(0, 0); }
            //    DependencyObject child = element.ItemContainerGenerator.ContainerFromIndex(index);
            //    Visual visual = FindVisualChild<Visual>(child);
            //    return visual.PointToScreen(new Point(10, 10)); // not working
            //});
        }

        public static T FindVisualChild<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        return (T)child;
                    }

                    T childItem = FindVisualChild<T>(child);
                    if (childItem != null) { return childItem; }
                }
            }
            return null;
        }

        public Point GetPosition<T>(string name)
        {
            return BlockUntilSTAThread(() =>
            {
                Control element = GetElement<T>(name) as Control;
                if (element == null) { return new Point(0, 0); }
                return element.PointToScreen(new Point(element.ActualWidth / 2, element.ActualHeight / 2));
            });
        }

        public void PressKey<T>(string name, Key key)
        {
            BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                PresentationSource source = PresentationSource.FromVisual(element);
                RoutedEvent routedEvent = Keyboard.KeyDownEvent; // Event to send

                element.RaiseEvent(
                  new KeyEventArgs(
                    Keyboard.PrimaryDevice,
                    source,
                    0,
                    key)
                  {
                      RoutedEvent = routedEvent
                  });
            });
        }

        public void TypeUsingKeyboard<T>(string name, string text)
        {
            BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                RoutedEvent routedEvent = TextCompositionManager.TextInputEvent;

                element.RaiseEvent(new TextCompositionEventArgs(
                    InputManager.Current.PrimaryKeyboardDevice,
                    new TextComposition(InputManager.Current, element, text))
                    {
                        RoutedEvent = routedEvent
                    });
            });
        }

        public bool IsType<T>() => BlockUntilSTAThread(() => window is T);

        public void Show()
        {
            BlockUntilSTAThread(window.Show);
        }

        public bool? ShowDialog()
        {
            return BlockUntilSTAThread(window.ShowDialog);
        }

        public void Close()
        {
            BlockUntilSTAThread(window.Close);
        }

        private UIElement GetElement<T>(string name)
        {
            return typeof(T).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(window) as UIElement;
        }

        private void BlockUntilSTAThread(Action action)
        {
            BlockUntilSTAThread<object>(() =>
            {
                action();
                return null;
            });
        }

        private T BlockUntilSTAThread<T>(Func<T> action)
        {
            if (window == null) { return default(T); }
            if (!window.Dispatcher.CheckAccess())
            {
                T result = default(T);
                window.Dispatcher.Invoke((Action)(() =>
                {
                    result = action();
                }));
                return result;
            }
            return action();
        }
    }
}
