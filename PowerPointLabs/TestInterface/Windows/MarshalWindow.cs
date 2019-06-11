using System;
using System.Reflection;
using System.Windows;

namespace TestInterface.Windows
{
    // TODO: Move into PowerpointLabs
    public class MarshalWindow : MarshalByRefObject
    {
        private readonly Window Window;

        public string Title => BlockUntilSTAThread(() => Window.Title);

        private MarshalWindow(Window w)
        {
            Window = w;
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
                UIElement element = typeof(T).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(Window) as UIElement;
                element.RaiseEvent(args);
            });
        }

        public bool Focus<T>(string name)
        {
            return BlockUntilSTAThread(() =>
            {
                UIElement element = typeof(T).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(Window) as UIElement;
                element.Focusable = true;
                return element.Focus();
            });
        }

        public bool IsType<T>() => BlockUntilSTAThread(() => Window is T);

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
            if (Window == null) { return default(T); }
            if (!Window.Dispatcher.CheckAccess())
            {
                T result = default(T);
                Window.Dispatcher.Invoke((Action)(() => {
                    result = action();
                }));
                return result;
            }
            return action();
        }

        public void Show()
        {
            BlockUntilSTAThread(Window.Show);
        }

        public bool? ShowDialog()
        {
            return BlockUntilSTAThread(Window.ShowDialog);
        }

        public void Close()
        {
            BlockUntilSTAThread(Window.Close);
        }

    }
}
