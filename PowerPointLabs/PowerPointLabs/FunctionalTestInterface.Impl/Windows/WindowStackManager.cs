using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;
using TestInterface.Windows;

namespace PowerPointLabs.FunctionalTestInterface.Windows
{
    [Serializable]
    public class WindowStackManager : MarshalByRefObject, IWindowStackManager
    {
        private Stack<IMarshalWindow> windowStack = new Stack<IMarshalWindow>();

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

        public IMarshalWindow Push(IntPtr handle)
        {
            Window window = GetWindow(handle);
            IMarshalWindow marshalWindow = MarshalWindow.CreateInstance(window);
            Push(marshalWindow);
            return marshalWindow;
        }

        public void Push(IMarshalWindow marshalWindow)
        {
            windowStack.Push(marshalWindow);
        }

        public IMarshalWindow Peek()
        {
            return (windowStack.Count == 0) ? null : windowStack.Peek();
        }

        public void Pop(bool close = true)
        {
            IMarshalWindow w = windowStack.Pop();
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
