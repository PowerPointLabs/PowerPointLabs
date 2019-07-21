using System;

namespace TestInterface.Windows
{
    public interface IWindowStackManager
    {
        void Setup();
        void Teardown();
        IMarshalWindow Push(IntPtr handle);
        void Push(IMarshalWindow marshalWindow);
        IMarshalWindow Peek();
        void Pop(bool close = true);
    }
}
