using System;
using System.Windows.Forms;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    /// <summary>
    /// 
    /// Must use this UI Thread Executor to
    /// do UI-related stuff:
    /// e.g. open Colors Lab or Shapes Lab pane.
    /// Otherwise, modification from FT thread won't
    /// succeed.
    /// 
    /// </summary>
    class UIThreadExecutor : Control
    {
        private static UIThreadExecutor _instance;
        private static bool _isDisposed = false;

        private UIThreadExecutor() {}

        public static void Init()
        {
            if (_instance == null)
            {
                _instance = new UIThreadExecutor {Visible = false};
                if (!_instance.IsHandleCreated)
                {
                    _instance.CreateHandle();
                }
            }
        }

        public static void TearDown()
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;
            _instance.Dispose(true);
        }

        public static void Execute(Action action)
        {
            _instance.Invoke(action);
        }
    }
}
