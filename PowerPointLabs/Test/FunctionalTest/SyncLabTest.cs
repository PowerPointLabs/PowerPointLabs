using System;
using System.Drawing;
using System.Windows.Forms;
using TestInterface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;
using Point = System.Drawing.Point;
using System.Collections.Generic;

namespace Test.FunctionalTest
{
    [TestClass]
    public class SyncLabTest : BaseFunctionalTest
    {
        private const int OriginalShapesSlideNo = 4;
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
        }

        private void TestCopyDialog(ISyncLabController syncLab)
        {
            var actualSlide = PpOperations.SelectSlide(4);

            // No item selected
            MessageBoxUtil.ExpectMessageBoxWillPopUp("", "Please select one item to copy.", syncLab.OpenCopyDialog, "Ok");

            // 2 item selected
            List<String> shapes = new List<string> { Oval, RotatedArrow };
            PpOperations.SelectShapes(shapes);
            MessageBoxUtil.ExpectMessageBoxWillPopUp("", "Please select one item to copy.", syncLab.OpenCopyDialog, "Ok");

            // copy successful
            PpOperations.SelectShape(CopyFromShape);
            syncLab.Copy();

        }

        # region Helper methods
        // mouse drag & drop from Control to Shape to apply color
        private void ApplyColor(Control from, Shape to)
        {
            var startPt = from.PointToScreen(new Point(from.Width/2, from.Height/2));
            var endPt = new Point(
                PpOperations.PointsToScreenPixelsX(to.Left + to.Width/2),
                PpOperations.PointsToScreenPixelsY(to.Top + to.Height/2));
            DragAndDrop(startPt, endPt);
        }

        // mouse drag & drop from control to another control to apply color
        private void ApplyColor(Control from, Control to)
        {
            var startPt = from.PointToScreen(new Point(from.Width / 2, from.Height / 2));
            var endPt = to.PointToScreen(new Point(to.Width / 2, to.Height / 2));
            DragAndDrop(startPt, endPt);
        }

        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown(startPt.X, startPt.Y);
            MouseUtil.SendMouseUp(endPt.X, endPt.Y);
        }

        private void Click(Control target)
        {
            var pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
            MouseUtil.SendMouseLeftClick(pt.X, pt.Y);
        }

        private static void AssertEqual(Color exp, Color actual)
        {
            // dont assert Alpha
            Assert.IsTrue(IsAlmostSame(exp.R, actual.R), "diff color R, exp {0}, actual {1}", exp.R, actual.R);
            Assert.IsTrue(IsAlmostSame(exp.G, actual.G),"diff color G, exp {0}, actual {1}", exp.G, actual.G);
            Assert.IsTrue(IsAlmostSame(exp.B, actual.B), "diff color B, exp {0}, actual {1}", exp.B, actual.B);
        }

        private static bool IsAlmostSame(byte a, byte b)
        {
            return Math.Abs(a - b) < 3;
        }
        # endregion
    }
}
