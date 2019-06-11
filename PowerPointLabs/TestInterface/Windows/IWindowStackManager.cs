using System;

namespace TestInterface.Windows
{
    public interface IWindowStackManager
    {
        void Setup();
        void Teardown();
        MarshalWindow Push(IntPtr handle);
        void Push(MarshalWindow marshalWindow);
        MarshalWindow Peek();
        void Pop(bool close = true);
    }
}
