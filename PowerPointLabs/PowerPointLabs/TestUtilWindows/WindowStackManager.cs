using System;
using System.Collections.Generic;
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

        public IMarshalWPF Push<T>(int handle, string title) where T : DispatcherObject
        {
            Window window = GetWindow(new IntPtr(handle));
            IMarshalWPF result = new MarshalWPF<T>(window, title);
            windowStack.Push(result);
            return result;
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

        public IMarshalWPF GetMarshalWPF<T>(IntPtr handle) where T : DispatcherObject
        {
            Window window = GetWindow(handle);
            return new MarshalWPF<T>(window, "placeholder");
        }

        private Window GetWindow(IntPtr handle)
        {
            HwndSource hwndSource = HwndSource.FromHwnd(handle);
            if (hwndSource == null) { return null; }
            return hwndSource.RootVisual as Window;
        }

    }
}
