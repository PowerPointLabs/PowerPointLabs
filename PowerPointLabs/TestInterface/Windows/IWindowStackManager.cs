using System;
using System.Windows.Threading;

namespace TestInterface.Windows
{
    public interface IWindowStackManager
    {
        void Setup();
        void Teardown();
        IMarshalWPF Push<T>(int window, string name) where T : DispatcherObject;
        IMarshalWPF Peek();
        void Pop(bool close = true);
        IMarshalWPF GetMarshalWPF<T>(IntPtr window) where T : DispatcherObject;
    }
}
