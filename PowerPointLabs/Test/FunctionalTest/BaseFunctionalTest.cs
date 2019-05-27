using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Base;
using Test.Util;

using TestInterface;

namespace Test.FunctionalTest
{
    [TestClass]
    public abstract class BaseFunctionalTest: TestAssemblyFixture
    {
        public TestContext TestContext { get; set; }

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
            PPLClipboard.Instance.Teardown();
            if (TestContext.CurrentTestOutcome != UnitTestOutcome.Passed)
            {
                if (!Directory.Exists(PathUtil.GetTestFailurePath()))
                {
                    Directory.CreateDirectory(PathUtil.GetTestFailurePath());
                }
                PpOperations.SavePresentationAs(
                    PathUtil.GetTestFailurePresentationPath(
                        TestContext.TestName + "_" +
                        GetTestingSlideName()));
            }
            PpOperations.ClosePresentation();
        }

        protected static void CheckIfClipboardIsRestored(Action action, int actualSlideNum, string shapeNameToBeCopied, int expSlideNum, string expShapeNameToDelete, string expCopiedShapeName)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideNum);
            ShapeRange shapeToBeCopied = PpOperations.SelectShape(shapeNameToBeCopied);
            Assert.AreEqual(1, shapeToBeCopied.Count);

            // Add this shape to clipboard
            PPLClipboard.Instance.LockClipboard();
            shapeToBeCopied.Copy();
            action();

            // Paste whatever in clipboard
            ShapeRange newShape = actualSlide.Shapes.Paste();
            PPLClipboard.Instance.ReleaseClipboard();

            // Check if pasted shape is the same as the shape added to clipboard originally
            Assert.AreEqual(shapeNameToBeCopied, newShape.Name);
            Assert.AreEqual(shapeToBeCopied.Count, newShape.Count);

            Slide expSlide = PpOperations.SelectSlide(expSlideNum);
            if (expShapeNameToDelete != "")
            {
                PpOperations.SelectShape(expShapeNameToDelete)[1].Delete();
            }

            //Set the pasted shape location because the location of the pasted shape is flaky
            Shape expCopied = PpOperations.SelectShape(expCopiedShapeName)[1];
            newShape.Top = expCopied.Top;
            newShape.Left = expCopied.Left;

            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void ConnectPpl()
        {
            const int waitTime = 3000;
            int retryCount = 5;
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
                    IPowerPointLabsFT ftInstance = (IPowerPointLabsFT) Activator.GetObject(typeof (IPowerPointLabsFT),
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
            Process pptProcess = new Process
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
            Process[] processes = Process.GetProcessesByName("POWERPNT");
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
            int retry = 5;
            while (Process.GetProcessesByName("POWERPNT").Length > 0
                && retry > 0)
            {
                retry--;
                ThreadUtil.WaitFor(1500);
            }

            if (Process.GetProcessesByName("POWERPNT").Length > 0)
            {
                foreach (Process process in Process.GetProcessesByName("POWERPNT"))
                {
                    process.Kill();
                }
            }
        }
    }
}
