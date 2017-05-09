using System;
using System.Drawing;
using System.Windows.Forms;
using TestInterface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using Point = System.Drawing.Point;
using System.Collections.Generic;
using System.Threading.Tasks;
using PowerPointLabs;

namespace Test.FunctionalTest
{
    [TestClass]
    public class SyncLabTest : BaseFunctionalTest
    {
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string CopyFromShape = "CopyFrom";


        protected override string GetTestingSlideName()
        {
            return "SyncLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_SyncLabTest()
        {
            var syncLab = PplFeatures.SyncLab;
            syncLab.OpenPane();

            TestCopyDialog(syncLab);
            TestPasteDialog(syncLab);
        }

        private void TestCopyDialog(ISyncLabController syncLab)
        {
            var actualSlide = PpOperations.SelectSlide(4);

            // no selection
            MessageBoxUtil.ExpectMessageBoxWillPopUp("", "Please select one item to copy.", syncLab.Copy, "Ok");

            // 2 item selected
            List<String> shapes = new List<string> { Oval, RotatedArrow };
            PpOperations.SelectShapes(shapes);
            MessageBoxUtil.ExpectMessageBoxWillPopUp("", "Please select one item to copy.", syncLab.Copy, "Ok");

            // copy successful
            PpOperations.SelectShape(CopyFromShape);
            new Task(() =>
            {
                ThreadUtil.WaitFor(1000);
                SendKeys.SendWait("{ENTER}");
            }).Start();
            syncLab.Copy();
        }

        private void TestPasteDialog(ISyncLabController syncLab)
        {
            PpOperations.SelectSlide(5);
            MessageBoxUtil.ExpectMessageBoxWillPopUp("", "Please select at least one item to apply.", () => syncLab.Sync(0), "Ok");
        }

        # region Helper methods
        // mouse drag & drop from Control to Shape to apply color
        private void FindDialog()
        {
            
        }
        # endregion
    }
}
