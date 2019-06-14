using System;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Test.Util
{
    class MessageBoxUtil
    {
        private const String MessageBoxIdentifier = "#32770";

        // retry count will be deprecating
        public static void ExpectMessageBoxWillPopUp(string title, string expContent, Action messageBoxTrigger,
            string buttonNameToClick = null, int retryCount = 5, int waitTime = 5000)
        {
            //Task expect = ExpectMessageBoxWillPopUp(title, expContent, buttonNameToClick, retryCount, waitTime);
            //messageBoxTrigger.Invoke();
            //VerifyExpectation(expect, retryCount, waitTime);
            IntPtr msgBoxHandle = WindowWatcher.Push(title, messageBoxTrigger, waitTime);
            AssertMessageboxContent(msgBoxHandle, retryCount, expContent);
            CloseMessageBox(msgBoxHandle, buttonNameToClick);
        }

        public static void ExpectMessageBoxWillNotPopUp(string title, string expContent, Action messageBoxTrigger,
            string buttonNameToClick = null, int retryCount = 5, int waitTime = 5000)
        {
            //Task expect = ExpectMessageBoxWillNotPopUp(title, expContent, buttonNameToClick, retryCount, waitTime);
            //messageBoxTrigger.Invoke();
            //VerifyExpectation(expect, retryCount, waitTime);
            try
            {
                IntPtr msgBoxHandle = WindowWatcher.Push(title, messageBoxTrigger, waitTime);
                CloseMessageBox(msgBoxHandle, buttonNameToClick);
            }
            catch (AssertFailedException)
            {
                // discard error
                return;
            }
            Assert.Fail("Message box should not open.");
        }

        private static void AssertMessageboxContent(IntPtr msgBoxHandle, int retryCount, string expContent)
        {
            IntPtr dlgHandle = IntPtr.Zero;
            for (int i = 0; i < 5; i++)
            {
                // this API is a bit iffy, not sure why unable to find sometimes
                dlgHandle = NativeUtil.GetDlgItem(msgBoxHandle, 0xFFFF);
                if (dlgHandle != IntPtr.Zero) { break; }
                ThreadUtil.WaitFor(1000);
            }
            Assert.AreNotEqual(IntPtr.Zero, dlgHandle, "Failed to find label in the messagebox.");

            const int nchars = 1024;
            StringBuilder actualContentBuilder = new StringBuilder(nchars);
            int isGetTextSuccessful = NativeUtil.GetWindowText(dlgHandle, actualContentBuilder, nchars);

            if (expContent != "{*}")
            {
                Assert.IsTrue(isGetTextSuccessful > 0, "Failed to get text in the label of messagebox.");
                Assert.AreEqual(expContent, actualContentBuilder.ToString(), true, "Different MessageBox content.");
            }
        }

        // This method must be called before PplFeatures,
        // otherwise, PplFeatures will block the test.
        private static Task ExpectMessageBoxWillPopUp(string title, string expContent,
            string buttonNameToClick, int retryCount, int waitTime)
        {
            // MessageBox in pptlabs will block the whole thread,
            // so multi-thread is needed here.
            Task taskToVerify = new Task(() =>
            {
                // try to find messagebox window
                IntPtr msgBoxHandle = IntPtr.Zero;
                while (msgBoxHandle == IntPtr.Zero && retryCount > 0)
                {
                    msgBoxHandle = NativeUtil.FindWindow(MessageBoxIdentifier, title);
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
                IntPtr dlgHandle = NativeUtil.GetDlgItem(msgBoxHandle, 0xFFFF);
                Assert.AreNotEqual(IntPtr.Zero, dlgHandle, "Failed to find label in the messagebox.");

                const int nchars = 1024;
                StringBuilder actualContentBuilder = new StringBuilder(nchars);
                int isGetTextSuccessful = NativeUtil.GetWindowText(dlgHandle, actualContentBuilder, nchars);

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

        // This method must be called before PplFeatures,
        // otherwise, PplFeatures will block the test.
        private static Task ExpectMessageBoxWillNotPopUp(string title, string expContent,
            string buttonNameToClick, int retryCount, int waitTime)
        {
            // MessageBox in pptlabs will block the whole thread,
            // so multi-thread is needed here.
            Task taskToVerify = new Task(() =>
            {
                // try to find messagebox window
                IntPtr msgBoxHandle = IntPtr.Zero;
                while (msgBoxHandle == IntPtr.Zero && retryCount > 0)
                {
                    msgBoxHandle = NativeUtil.FindWindow(MessageBoxIdentifier, title);
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
                if (msgBoxHandle != IntPtr.Zero)
                {
                    Assert.Fail("Message Box should not be open.");
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
            // don't end so fast..
            ThreadUtil.WaitFor(500);
        }

        private static void CloseMessageBox(IntPtr msgBoxHandle, string buttonName)
        {
            if (buttonName == null)
            {
                // Simple close message box
                NativeUtil.SetForegroundWindow(msgBoxHandle);
                NativeUtil.SendMessage(msgBoxHandle, 0x0112 /*WM_SYSCOMMAND*/, new IntPtr(0xF060 /*SC_CLOSE*/), IntPtr.Zero);
            }
            else
            {
                // This may be flaky.. if there're more than one windows pop up at the same time..
                // it will affect clicking the button
                IntPtr btnHandle = NativeUtil.FindWindowEx(msgBoxHandle, IntPtr.Zero, "Button", buttonName);
                Assert.AreNotEqual(IntPtr.Zero, btnHandle, "Failed to find button in the messagebox.");
                NativeUtil.SetForegroundWindow(msgBoxHandle);
                NativeUtil.SendMessage(btnHandle, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
                NativeUtil.SendMessage(btnHandle, 0x0202 /*left button up*/, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}