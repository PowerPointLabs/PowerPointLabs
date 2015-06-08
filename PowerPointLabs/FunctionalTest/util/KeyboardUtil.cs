using System.Windows.Forms;

namespace FunctionalTest.util
{
    class KeyboardUtil
    {
        public static void Copy()
        {
            NativeUtil.SetForegroundWindow(NativeUtil.FindWindow("PPTFrameClass", null));
            SendKeys.SendWait("^c");
        }

        public static void Paste()
        {
            NativeUtil.SetForegroundWindow(NativeUtil.FindWindow("PPTFrameClass", null));
            SendKeys.SendWait("^v");
        }
    }
}
