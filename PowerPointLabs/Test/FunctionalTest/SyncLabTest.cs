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

            TestSync(syncLab);
            TestErrorDialogs(syncLab);
        }

        private void TestErrorDialogs(ISyncLabController syncLab)
        {
            PpOperations.SelectSlide(4);

            // no selection
            MessageBoxUtil.ExpectMessageBoxWillPopUp(TextCollection.SyncLabErrorDialogTitle,
                "Please select one item to copy.", syncLab.Copy, "Ok");

            // 2 item selected
            List<String> shapes = new List<string> { Oval, RotatedArrow };
            PpOperations.SelectShapes(shapes);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(TextCollection.SyncLabErrorDialogTitle,
                "Please select one item to copy.", syncLab.Copy, "Ok");

            // copy blank item for the paste error dialog test
            PpOperations.SelectShape(CopyFromShape);    
            syncLab.Copy();
            syncLab.DialogClickOk();

            PpOperations.SelectSlide(5);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(TextCollection.SyncLabErrorDialogTitle,
                "Please select at least one item to apply.", () => syncLab.Sync(0), "Ok");
        }

        private void TestSync(ISyncLabController syncLab)
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShape(CopyFromShape);

            syncLab.Copy();
            syncLab.DialogSelectItem(3, 2);
            syncLab.DialogClickOk();

            PpOperations.SelectShape(UnrotatedRectangle);
            syncLab.Sync(0);

            var actualSlide = PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShape(UnrotatedRectangle)[1];
            var expectedSlide = PpOperations.SelectSlide(5);
            var expectedShape = PpOperations.SelectShape(UnrotatedRectangle)[1];
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameShape(expectedShape, actualShape);
        }
    }
}
