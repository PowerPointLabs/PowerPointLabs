using System;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest.util
{
    class MessageBoxUtil
    {
        public static void ExpectMessageBoxWillPopUp(string title, string expContent, Action messageBoxTrigger, 
            string buttonNameToClick = null, int retryCount = 5, int waitTime = 1000)
        {
            Task expect = ExpectMessageBoxWillPopUp(title, expContent, buttonNameToClick, retryCount, waitTime);
            messageBoxTrigger.Invoke();
            VerifyExpectation(expect, retryCount, waitTime);
        }

        // This method must be called before PplFeatures,
        // otherwise, PplFeatures will block the test.
        private static Task ExpectMessageBoxWillPopUp(string title, string expContent,
            string buttonNameToClick, int retryCount, int waitTime)
        {
            // MessageBox in pptlabs will block the whole thread,
            // so multi-thread is needed here.
            var taskToVerify = new Task(() =>
            {
                // try to find messagebox window
                var msgBoxHandle = IntPtr.Zero;
                while (msgBoxHandle == IntPtr.Zero && retryCount > 0)
                {
                    msgBoxHandle = NativeUtil.FindWindow("#32770", title);
                    if (msgBoxHandle == IntPtr.Zero)
                    {
                        ThreadUtil.WaitFor(waitTime);
                        retryCount--;
                    }
                    else
                    {
                        break;
                    }
                }
                if (msgBoxHandle == IntPtr.Zero)
                {
                    Assert.Fail("Failed to find message box.");
                }

                // try to find text label in the message box
                var dlgHandle = NativeUtil.GetDlgItem(msgBoxHandle, 0xFFFF);
                Assert.AreNotEqual(IntPtr.Zero, dlgHandle, "Failed to find label in the messagebox.");

                const int nchars = 1024;
                var actualContentBuilder = new StringBuilder(nchars);
                var isGetTextSuccessful = NativeUtil.GetWindowText(dlgHandle, actualContentBuilder, nchars);

                // close the message box, otherwise it will block the test
                CloseMessageBox(msgBoxHandle, buttonNameToClick);

                if (expContent != "{*}")
                {
                    Assert.IsTrue(isGetTextSuccessful > 0, "Failed to get text in the label of messagebox.");
                    Assert.AreEqual(expContent, actualContentBuilder.ToString(), true, "Different MessageBox content.");
                }
            });
            taskToVerify.Start();
            return taskToVerify;
        }

        private static void VerifyExpectation(Task taskToVerify, int retryCount, int waitTime)
        {
            // wait for task to finish
            while (taskToVerify.Status == TaskStatus.Running && retryCount > 0)
            {
                ThreadUtil.WaitFor(waitTime);
                retryCount--;
            }
            // assert no exception during task's execution
            if (taskToVerify.Exception != null)
            {
                Assert.AreEqual(null, taskToVerify.Exception, "Failed to verify expectation. Exception: {0}",
                    taskToVerify.Exception.Message);
            }
        }

        private static void CloseMessageBox(IntPtr msgBoxHandle, string buttonName)
        {
            if (buttonName == null)
            {
                // Simple close message box
                NativeUtil.SendMessage(msgBoxHandle, 0x0112 /*WM_SYSCOMMAND*/, new IntPtr(0xF060 /*SC_CLOSE*/), IntPtr.Zero);
            }
            else
            {
                // This may be flaky.. if there're more than one windows pop up at the same time..
                // it will affect clicking the button
                var btnHandle = NativeUtil.FindWindowEx(msgBoxHandle, IntPtr.Zero, "Button", buttonName);
                Assert.AreNotEqual(IntPtr.Zero, btnHandle, "Failed to find button in the messagebox.");
                NativeUtil.SetForegroundWindow(msgBoxHandle);
                NativeUtil.SendMessage(btnHandle, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
                NativeUtil.SendMessage(btnHandle, 0x0202 /*left button up*/, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}
