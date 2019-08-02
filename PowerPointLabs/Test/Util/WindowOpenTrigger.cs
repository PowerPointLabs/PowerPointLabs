using System;
using System.Threading;

namespace Test.Util
{
    class WindowOpenTrigger : ManualResetEventSlim
    {
        public IntPtr resultingWindow;
        public int[] runtimeId;
        public WindowOpenTrigger(bool initialState) : base(initialState)
        {

        }
    }
}
