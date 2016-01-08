using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting;
using System.Windows.Forms;
using FunctionalTest.util;
using FunctionalTestInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public abstract class BaseFunctionalTest
    {
        public TestContext TestContext { get; set; }

        private static int numberOfFailedTest = 0;

        // prefix legend:
        // pp - PowerPoint
        // ppl - PowerPointLabs
        protected static IPowerPointLabsFeatures PplFeatures;
        protected static IPowerPointOperations PpOperations;

        // To be implemented by downstream testing classes,
        // specify the name for the testing slide.
        // It is assumed that the testing slides reside
        // in "doc/test" folder.
        protected abstract string GetTestingSlideName();

        // To be override by some test case to use new
        // PowerPoint application instance for FT
        protected virtual bool IsUseNewPpInstance()
        {
            return false;
        }

        [AssemblyInitialize]
        public static void AssemblySetup(TestContext context)
        {
            var folderToClean = new DirectoryInfo(PathUtil.GetTestFailurePath());
            if (folderToClean.Exists)
            {
                foreach (var file in folderToClean.GetFiles())
                {
                    file.Delete();
                }
            }
        }

        [TestInitialize]
        public void Setup()
        {
            if (IsUseNewPpInstance())
            {
                CloseActivePpInstance();
            }

            OpenSlideForTest(GetTestingSlideName());

            ConnectPpl();
        }

        [TestCleanup]
        public void TearDown()
        {
            if (TestContext.CurrentTestOutcome != UnitTestOutcome.Passed)
            {
                numberOfFailedTest++;
                if (!Directory.Exists(PathUtil.GetTestFailurePath()))
                {
                    Directory.CreateDirectory(PathUtil.GetTestFailurePath());
                }
                PpOperations.SavePresentationAs(
                    PathUtil.GetTestFailurePresentationPath(
                        GetTestingSlideName()));
            }
            PpOperations.ClosePresentation();
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            if (numberOfFailedTest != 0)
            {
                MessageBox.Show("Failed cases found. Please check failed slides in the folder 'doc\\test\\TestFailed'.\n" +
                                "Please submit an issue with the failed slides if the problem persists.");
            }
            else
            {
                MessageBox.Show("Pass!");
            }
        }

        private void ConnectPpl()
        {
            const int waitTime = 3000;
            var retryCount = 5;
            while (retryCount > 0)
            {
                // if already connected, break
                if (PplFeatures != null && PpOperations != null)
                {
                    break;
                }
                // otherwise keep trying to connect for some times
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
                    ThreadUtil.WaitFor(waitTime);
                }
            }
            if (PplFeatures == null || PpOperations == null)
            {
                Assert.Fail("Failed to connect to PowerPointLabs add-in. You can try to increase retryCount.");
            }

            PpOperations.EnterFunctionalTest();

            // activate the thread of presentation window
            ThreadUtil.WaitFor(1500);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "PowerPointLabs FT", "{*}",
                PpOperations.ActivatePresentation);
        }

        private void OpenSlideForTest(String slideName)
        {
            var pptProcess = new Process
            {
                StartInfo =
                {
                    FileName = slideName, 
                    WorkingDirectory = PathUtil.GetDocTestPath()
                }
            };
            pptProcess.Start();
        }

        private void CloseActivePpInstance()
        {
            var processes = Process.GetProcessesByName("POWERPNT");
            if (processes.Length > 0)
            {
                foreach (Process p in processes)
                {
                    p.CloseMainWindow();
                }
            }
            WaitForPpInstanceToClose();
            PpOperations = null;
            PplFeatures = null;
        }

        private void WaitForPpInstanceToClose()
        {
            var retry = 5;
            while (Process.GetProcessesByName("POWERPNT").Length > 0
                && retry > 0)
            {
                retry--;
                ThreadUtil.WaitFor(1500);
            }

            if (Process.GetProcessesByName("POWERPNT").Length > 0)
            {
                foreach (var process in Process.GetProcessesByName("POWERPNT"))
                {
                    process.Kill();
                }
            }
        }
    }
}
