using System;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Test.Util
{
    class DialogUtil
    {
        public static void WaitForDialogBox(Action openDialogAction, string lpClassName, string lpWindowName, int timeLimit = 5000)
        {
            Task task = new Task(() => openDialogAction());
            task.Start();

            int pollCount = 10;
            int retryInterval = timeLimit/pollCount;
            for (int i = 0; i <= pollCount; ++i)
            {
                IntPtr spotlightDialog = NativeUtil.FindWindow(null, lpWindowName);
                if (spotlightDialog != IntPtr.Zero) return;
                ThreadUtil.WaitFor(retryInterval);
            }
            Assert.Fail("Wait for dialog box timed out");
        }

        public static void CloseDialogBox(IntPtr dialogBoxHandle, string buttonName)
        {
            IntPtr btnHandle = NativeUtil.FindWindowEx(dialogBoxHandle, IntPtr.Zero, null, buttonName);
            Assert.AreNotEqual(IntPtr.Zero, btnHandle, "Failed to find button in the dialog box.");
            NativeUtil.SetForegroundWindow(dialogBoxHandle);
            NativeUtil.SendMessage(btnHandle, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
            NativeUtil.SendMessage(btnHandle, 0x0202 /*left button up*/, IntPtr.Zero, IntPtr.Zero);
        }
    }
}
