using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using FunctionalTestInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    public abstract class BaseFunctionalTest
    {
        protected static IPowerPointLabsFeatures PplFeatures;
        protected static IPowerPointOperations PpOperations;
        protected static Process proc;

        protected abstract String GetSlideName();

        protected void WaitFor(int time)
        {
            var waitedFor = 0;
            while (waitedFor < time)
            {
                Thread.Sleep(time);
                waitedFor += time;
            }
        }

        [TestInitialize]
        public void Setup()
        {
            OpenSlideForTest(GetSlideName());
            WaitFor(1500);
            var FtInstance = (IPowerPointLabsFT) Activator.GetObject(typeof(IPowerPointLabsFT),
                        "ipc://PowerPointLabsFT/PowerPointLabsFT");
            PplFeatures = FtInstance.GetFeatures();
            PpOperations = FtInstance.GetOperations();
            PpOperations.EnterFunctionalTest();
            Assert.IsTrue(PpOperations.IsInFunctionalTest());
        }

        [TestCleanup]
        public void TearDown()
        {
            Assert.IsTrue(PpOperations.IsInFunctionalTest());
            if (proc != null)
            {
                proc.CloseMainWindow();
            }
        }

        public void OpenSlideForTest(String slideName)
        {
            //To get the location the assembly normally resides on disk or the install directory
            var path = new Uri(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase))
                .LocalPath;
            var parPath = GetParentFolder(GetParentFolder(GetParentFolder(GetParentFolder(path))));
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
            proc = pptProcess;
        }

        private String GetParentFolder(String path)
        {
            return Directory.GetParent(path).FullName;
        }
    }
}
