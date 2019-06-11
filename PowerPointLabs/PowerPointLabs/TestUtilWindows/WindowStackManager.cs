using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;
using TestInterface.Windows;

namespace Test.Util
{
    [Serializable]
    public class WindowStackManager : MarshalByRefObject, IWindowStackManager
    {
        private Stack<MarshalWindow> windowStack = new Stack<MarshalWindow>();

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

        public MarshalWindow Push(IntPtr handle)
        {
            Window window = GetWindow(handle);
            MarshalWindow marshalWindow = MarshalWindow.CreateInstance(window);
            Push(marshalWindow);
            return marshalWindow;
        }

        public void Push(MarshalWindow marshalWindow)
        {
            windowStack.Push(marshalWindow);
        }

        public MarshalWindow Peek()
        {
            return (windowStack.Count == 0) ? null : windowStack.Peek();
        }

        public void Pop(bool close = true)
        {
            MarshalWindow w = windowStack.Pop();
            if (close)
            {
                w?.Close();
            }
        }

        private Window GetWindow(IntPtr handle)
        {
            HwndSource hwndSource = HwndSource.FromHwnd(handle);
            if (hwndSource == null) { return null; }
            return hwndSource.RootVisual as Window;
        }

    }
}
