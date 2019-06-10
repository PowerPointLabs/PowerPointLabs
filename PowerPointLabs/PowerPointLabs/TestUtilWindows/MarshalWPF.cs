using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using TestInterface.Windows;

namespace Test.Util
{
    /// <summary>
    /// A class that allows sending <seealso cref="RoutedEventArgs"/> to elements exposed as first layer fields.
    /// </summary>
    /// <typeparam name="RefObject">The type of WPF element</typeparam>
    [Serializable]
    public class MarshalWPF<RefObject> : MarshalByRefObject, IMarshalWPF where RefObject : DispatcherObject
    {
        private Window obj;
        private ManualResetEventSlim canExecute = new ManualResetEventSlim(false);
        public string Title { get; private set; } // the compiler complains for IMarshalWPF.Title
        public Type Type => typeof(RefObject);

        public MarshalWPF(Window obj, string title)
        {
            this.obj = obj;
            this.Title = title;
        }

        public void RaiseEvent(string name, RoutedEventArgs args)
        {
            UIElement element = typeof(RefObject).GetField(name).GetValue(obj) as UIElement;
            element.RaiseEvent(args);
        }

        public bool Focus(string name)
        {
            UIElement element = typeof(RefObject).GetField(name).GetValue(obj) as UIElement;
            return element.Focus();
        }

        public void Close()
        {
            if (obj is Window)
            {
                (obj as Window)?.Close();
            }
        }

        private T BlockUntilSTAThread<T>(Func<T> action)
        {
            if (!obj.Dispatcher.CheckAccess())
            {
                T result = default(T);
                canExecute.Reset();
                obj.Dispatcher.Invoke((Action)(() =>
                {
                    result = action();
                    canExecute.Set();
                }));
                canExecute.Wait();
                return result;
            }
            return action();
        }
    }
}
