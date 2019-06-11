using System;
using System.Threading;

namespace Test.Util.Windows
{
    class WindowOpenTrigger : ManualResetEventSlim
    {
        public IntPtr resultingWindow;
        public string name;
        public WindowOpenTrigger(bool initialState) : base(initialState)
        {

        }
    }
}
