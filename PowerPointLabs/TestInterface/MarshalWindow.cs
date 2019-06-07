using System;
using System.Windows;

namespace TestInterface
{
    public class MarshalWindow : MarshalByRefObject
    {
        public Window Window;
        public MarshalWindow(Window w)
        {
            Window = w;
        }
    }
}
