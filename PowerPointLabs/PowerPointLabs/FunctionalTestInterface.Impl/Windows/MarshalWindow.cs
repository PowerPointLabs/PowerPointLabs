using System;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using TestInterface.Windows;

namespace PowerPointLabs.FunctionalTestInterface.Windows
{
    /// <summary>
    /// A class used to Marshal a window reference across 2 applications.
    /// Supports raising WPF events on first level field elements
    /// </summary>
    public class MarshalWindow : MarshalByRefObject, IMarshalWindow
    {
        private readonly Window window;

        public string Title => BlockUntilSTAThread(() => window.Title);

        private MarshalWindow(Window w)
        {
            window = w;
        }

        /// <summary>
        /// Failable construction of MarshalWindow, fails if window is null.
        /// </summary>
        /// <param name="w">Window to be marshalled</param>
        /// <returns>Marshalled window</returns>
        public static MarshalWindow CreateInstance(Window w)
        {
            if (w == null) { return null; }
            return new MarshalWindow(w);
        }

        /// <summary>
        /// Raises a WPF event to element with the specified field name.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <param name="args">WPF event arguments</param>
        public void RaiseEvent<T>(string name, RoutedEventArgs args)
        {
            BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                element.RaiseEvent(args);
            });
        }

        /// <summary>
        /// Clicks on the UI element.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
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

        /// <summary>
        /// Attempts to focus on the soecified UI element.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <returns>Whether focus is successful</returns>
        public bool Focus<T>(string name)
        {
            return BlockUntilSTAThread(() =>
            {
                UIElement element = GetElement<T>(name);
                element.Focusable = true;
                return element.Focus();
            });
        }

        /// <summary>
        /// Selects all the text in the specified UI element.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        public void SelectAll<T>(string name)
        {
            BlockUntilSTAThread(() =>
            {
                TextBox element = GetElement<T>(name) as TextBox;
                element?.SelectAll();
            });
        }

        /// <summary>
        /// Gets the position of an element in a list using its index.
        /// </summary>
        /// <remarks>
        /// Assumes that the list is a vertical list, with element dimensions same as its parent
        /// </remarks>
        /// <typeparam name="T">Type of list</typeparam>
        /// <param name="name">Name of list</param>
        /// <param name="index">Index of UI element</param>
        /// <returns>Position of UI element</returns>
        public Point GetListElementPosition<T>(string name, int index)
        {
            return BlockUntilSTAThread(() =>
            {
                Control element = GetElement<T>(name) as Control;
                if (element == null) { return new Point(0, 0); }
                int factor = 3 + 2 * index;
                return element.PointToScreen(new Point(element.ActualWidth / 2, element.ActualHeight * factor / 2));
            });
        }

        /// <summary>
        /// Gets the position of an UI element of specified name and type.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <returns>Position of UI element</returns>
        public Point GetPosition<T>(string name)
        {
            return BlockUntilSTAThread(() =>
            {
                Control element = GetElement<T>(name) as Control;
                if (element == null) { return new Point(0, 0); }
                return element.PointToScreen(new Point(element.ActualWidth / 2, element.ActualHeight / 2));
            });
        }

        /// <summary>
        /// Presses a specified <seealso cref="Key"/> into an UI element with specified name.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <param name="key"><seealso cref="Key"/> pressed</param>
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

        /// <summary>
        /// Types text into a UI element with a specified name.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <param name="text">Text to be keyed in</param>
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

        /// <summary>
        /// Checks if a <seealso cref="CheckBox"/> is checked.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <returns>Whether the element is checked</returns>
        public bool? IsChecked<T>(string name) => BlockUntilSTAThread(() => (GetElement<T>(name) as CheckBox).IsChecked);

        /// <summary>
        /// Checks if the underlying type of the window is T.
        /// </summary>
        /// <typeparam name="T">Type to be checked against</typeparam>
        /// <returns></returns>
        public bool IsType<T>() => BlockUntilSTAThread(() => window is T);

        /// <summary>
        /// Same as <seealso cref="Window.Show"/>
        /// </summary>
        public void Show()
        {
            BlockUntilSTAThread(window.Show);
        }

        /// <summary>
        /// Same as <seealso cref="Window.ShowDialog"/>
        /// </summary>
        public bool? ShowDialog()
        {
            return BlockUntilSTAThread(window.ShowDialog);
        }

        /// <summary>
        /// Same as <seealso cref="Window.Close"/>
        /// </summary>
        public void Close()
        {
            BlockUntilSTAThread(window.Close);
        }

        /// <summary>
        /// Retrieves a UI element with a specified name and type.
        /// </summary>
        /// <typeparam name="T">Type of UI element</typeparam>
        /// <param name="name">Name of UI element</param>
        /// <returns>UI Element</returns>
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

        /// <summary>
        /// Runs the code on the same thread that created it.
        /// </summary>
        /// <param name="action">Code to be ran</param>
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
