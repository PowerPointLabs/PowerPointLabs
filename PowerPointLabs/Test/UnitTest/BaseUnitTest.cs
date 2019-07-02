using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.Utils;

using Test.Base;
using Test.Util;

using TestInterface;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace Test.UnitTest
{
    [TestClass]
    public abstract class BaseUnitTest: TestAssemblyFixture
    {
        public TestContext TestContext { get; set; }

        protected IPowerPointOperations PpOperations;

        protected PowerPoint.Application App;

        protected PowerPoint.Presentation Pres;

        // To be implemented by downstream testing classes,
        // specify the name for the testing slide.
        // It is assumed that the testing slides reside
        // in "doc/test" folder.
        protected abstract string GetTestingSlideName();

        [TestInitialize]
        public void Setup()
        {
            CultureUtil.SetDefaultCulture(CultureInfo.GetCultureInfo("en-US"));
            try
            {
                App = new PowerPoint.Application();
            }
            catch (COMException)
            {
                // in case a warm-up is needed
                App = new PowerPoint.Application();
            }
            Pres = App.Presentations.Open(
                PathUtil.GetDocTestPath() + GetTestingSlideName(),
                WithWindow: MsoTriState.msoFalse);
            PpOperations = new UnitTestPpOperations(Pres, App);
        }

        protected void CheckIfClipboardIsRestored(Action action, int actualSlideNum, string shapeNameToBeCopied, int expSlideNum, string expShapeNameToDelete, string expCopiedShapeName)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideNum);
            ShapeRange shapeToBeCopied = PpOperations.SelectShape(shapeNameToBeCopied);
            Assert.AreEqual(1, shapeToBeCopied.Count);

            // Add this shape to clipboard
            shapeToBeCopied.Copy();
            action();

            // Paste whatever in clipboard
            ShapeRange newShape = actualSlide.Shapes.Paste();

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

            Util.SlideUtil.IsSameLooking(expSlide, actualSlide);
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
