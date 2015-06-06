using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting;
using FunctionalTest.util;
using FunctionalTestInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public abstract class BaseFunctionalTest
    {
        // prefix legend:
        // pp - PowerPoint
        // ppl - PowerPointLabs
        protected static IPowerPointLabsFeatures PplFeatures;
        protected static IPowerPointOperations PpOperations;
        protected static Process PpProcess;

        // To be implemented by downstream testing classes,
        // specify the name for the testing slide.
        // It is assumed that the testing slides reside
        // in "doc/test" folder.
        protected abstract String GetTestingSlideName();

        // Can be overrided by downstream testing classes,
        // e.g. increase the count.
        // For now, this retry count is used when try to connect
        // pptlabs add-in.
        protected virtual int GetRetryCount()
        {
            return 5;
        }

        protected virtual int GetWaitTime()
        {
            return 3000;
        }

        [TestInitialize]
        public void Setup()
        {
            AssertNoPpProcessRunning();

            OpenSlideForTest(GetTestingSlideName());

            ConnectPplWithRetryCount(GetRetryCount());

            PpOperations.EnterFunctionalTest();
            Assert.IsTrue(PpOperations.IsInFunctionalTest());
        }

        [TestCleanup]
        public void TearDown()
        {
            Assert.IsTrue(PpOperations.IsInFunctionalTest());
            if (PpProcess != null)
            {
                PpProcess.CloseMainWindow();
            }
            // wait for process to quit entirely
            AssertNoPpProcessRunningWithRetryCount(GetRetryCount());
        }

        private void ConnectPplWithRetryCount(int retryCount)
        {
            while (retryCount > 0)
            {
                try
                {
                    var ftInstance = (IPowerPointLabsFT) Activator.GetObject(typeof (IPowerPointLabsFT),
                        "ipc://PowerPointLabsFT/PowerPointLabsFT");
                    PplFeatures = ftInstance.GetFeatures();
                    PpOperations = ftInstance.GetOperations();
                    break;
                }
                catch (RemotingException)
                {
                    retryCount--;
                    ThreadUtil.WaitFor(GetWaitTime());
                }
            }
            if (retryCount == 0 || PplFeatures == null || PpOperations == null)
            {
                Assert.Fail("Failed to connect to PowerPointLabs add-in. You can try to increase retryCount.");
            }
        }

        private void AssertNoPpProcessRunningWithRetryCount(int retryCount)
        {
            while (retryCount > 0)
            {
                var pname = Process.GetProcessesByName("POWERPNT");
                if (pname.Length > 0)
                {
                    retryCount--;
                    ThreadUtil.WaitFor(GetWaitTime());
                }
                else
                {
                    break;
                }
            }
            if (retryCount == 0)
            {
                Assert.Fail("PowerPoint process is still running after a long time.");
            }
        }

        // When a PowerPoint process is already running,
        // it will affect the TearDown process.
        private void AssertNoPpProcessRunning()
        {
            var pname = Process.GetProcessesByName("POWERPNT");
            if (pname.Length > 0)
            {
                Assert.Fail("PowerPoint is running, you need to close it to continue Functional Test.");
            }
        }

        private void OpenSlideForTest(String slideName)
        {
            //To get the location the assembly normally resides on disk or the install directory
            var path = new Uri(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase))
                .LocalPath;
            var parPath = PathUtil.GetParentFolder(path, 4);
            var testDocPath = Path.Combine(parPath, "doc\\test\\");
            Process pptProcess = new Process
            {
                StartInfo =
                {
                    FileName = slideName, 
                    WorkingDirectory = testDocPath
                }
            };
            pptProcess.Start();
            PpProcess = pptProcess;
        }
    }
}
