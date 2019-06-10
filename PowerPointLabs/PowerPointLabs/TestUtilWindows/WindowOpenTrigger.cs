using System.Threading;

namespace Test.Util.Windows
{
    class WindowOpenTrigger : ManualResetEventSlim
    {
        public int resultingWindow;
        public string name;
        public WindowOpenTrigger(bool initialState) : base(initialState)
        {

        }
    }
}
