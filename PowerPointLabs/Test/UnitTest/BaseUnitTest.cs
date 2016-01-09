using System.IO;
using TestInterface;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest
{
    [TestClass]
    public abstract class BaseUnitTest
    {
        protected static PowerPoint.Application App;

        public TestContext TestContext { get; set; }

        protected IPowerPointOperations PpOperations;

        protected PowerPoint.Presentation Pres;

        // To be implemented by downstream testing classes,
        // specify the name for the testing slide.
        // It is assumed that the testing slides reside
        // in "doc/test" folder.
        // NOTES:
        // If no test slide is needed, return null or empty 
        // string.
        protected abstract string GetTestingSlideName();

        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            App = new PowerPoint.Application();
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            App.Quit();
        }

        [TestInitialize]
        public void Setup()
        {
            Pres = App.Presentations.Open(
                PathUtil.GetDocTestPath() + GetTestingSlideName(),
                WithWindow: MsoTriState.msoFalse);
            PpOperations = new UnitTestPpOperations(Pres, App);
        }

        [TestCleanup]
        public void TearDown()
        {
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
    }
}
