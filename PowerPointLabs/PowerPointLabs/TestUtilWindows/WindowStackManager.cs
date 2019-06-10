using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using Test.Util.Windows;
using TestInterface.Windows;

namespace Test.Util
{
    [Serializable]
    public class WindowStackManager : IWindowStackManager
    {
        private Stack<IMarshalWPF> windowStack = new Stack<IMarshalWPF>();

        public void Setup()
        {

        }

        public void Teardown()
        {
            while (Peek() != null)
            {
                Pop();
            }
        }

        public IMarshalWPF Push<T>(Window window) where T : DispatcherObject
        {
            return BlockUntilSTAThread<IMarshalWPF>(window, () => PushSTAThread<T>(window));
        }

        public IMarshalWPF Peek()
        {
            return (windowStack.Count == 0) ? null : windowStack.Peek();
        }

        public void Pop(bool close = true)
        {
            IMarshalWPF w = windowStack.Pop();
            if (close)
            {
                w?.Close();
            }
        }

        private IMarshalWPF PushSTAThread<T>(Window window) where T : DispatcherObject
        {
            IMarshalWPF result = new MarshalWPF<T>(window, window.Title);
            windowStack.Push(result);
            return result;
        }

        private void BlockUntilSTAThread(Window window, Action action)
        {
            BlockUntilSTAThread<object>(window, () =>
            {
                action();
                return null;
            });
        }

        private T BlockUntilSTAThread<T>(Window window, Func<T> action)
        {
            if (!window.Dispatcher.CheckAccess())
            {
                T result = default(T);
                ManualResetEventSlim canExecute = new ManualResetEventSlim(false);
                window.Dispatcher.Invoke((Action)(() =>
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
