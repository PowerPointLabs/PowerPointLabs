using System.Threading;
using System.Windows;

namespace Test.Util.Windows
{
    class WindowOpenTrigger : ManualResetEventSlim
    {
        public Window resultingWindow;

        public WindowOpenTrigger(bool initialState) : base(initialState)
        {

        }
    }
}
