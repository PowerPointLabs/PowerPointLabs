using System;
using System.Diagnostics;
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

        // To be implemented by downstream testing classes,
        // specify the name for the testing slide.
        // It is assumed that the testing slides reside
        // in "doc/test" folder.
        protected abstract string GetTestingSlideName();

        [TestInitialize]
        public void Setup()
        {
            OpenSlideForTest(GetTestingSlideName());

            ConnectPpl();
        }

        [TestCleanup]
        public void TearDown()
        {
            PpOperations.ClosePresentation();
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
    }
}
