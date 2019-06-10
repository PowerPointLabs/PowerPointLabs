using System;
using System.Threading;
using System.Windows;

namespace TestInterface.Windows
{
    public class MarshalWindow : MarshalByRefObject
    {
        private readonly IWindow Window;
        private ManualResetEventSlim canExecute;

        public string Title => Window.Title;

        public MarshalWindow(Window w)
        {
            Window = new WindowExt(w);
            canExecute = new ManualResetEventSlim(false);
        }

        public MarshalWindow(IWindow w)
        {
            Window = w;
            canExecute = new ManualResetEventSlim(false);
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
            if (!Window.Dispatcher.CheckAccess())
            {
                T result = default(T);
                canExecute.Reset();
                Window.Dispatcher.Invoke((Action)(() => {
                    result = action();
                    canExecute.Set();
                }));
                canExecute.Wait();
                return result;
            }
            return action();
        }

        public bool IsType<T>()
        {
            return Window is T;
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
