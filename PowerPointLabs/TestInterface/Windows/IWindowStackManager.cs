using System;
using System.Windows;
using System.Windows.Threading;

namespace TestInterface.Windows
{
    public interface IWindowStackManager
    {
        void Setup();
        void Teardown();
        IMarshalWPF Push<T>(Window window) where T : DispatcherObject;
        IMarshalWPF Peek();
        void Pop(bool close = true);
    }
}
