using System.Windows.Forms;

namespace FunctionalTest.util
{
    class KeyboardUtil
    {
        public static void Copy()
        {
            ThreadUtil.WaitFor(250);
            NativeUtil.SetForegroundWindow(NativeUtil.FindWindow("PPTFrameClass", null));
            SendKeys.SendWait("^c");
            ThreadUtil.WaitFor(250);
        }

        public static void Paste()
        {
            ThreadUtil.WaitFor(250);
            NativeUtil.SetForegroundWindow(NativeUtil.FindWindow("PPTFrameClass", null));
            SendKeys.SendWait("^v");
            ThreadUtil.WaitFor(250);
        }
    }
}
