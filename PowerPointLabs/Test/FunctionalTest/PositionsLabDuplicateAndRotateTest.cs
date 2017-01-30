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
    public class PositionsLabDuplicateAndRotateTest : BaseFunctionalTest
    {
        List<string> _shapeNames;

        private const string Rectangle = "Rectangle";
        private const string Oval = "Oval";
        private const string Triangle = "Triangle";
        private const string Target = "Target";

        private const int OriginalShapesSlideNoTestOneFixed = 4;
        private const int OriginalShapesSlideNoTestOneDynamic = 7;
        private const int OriginalShapesSlideNoTestMultipleFixed = 10;
        private const int OriginalShapesSlideNoTestMultipleDynamic = 13;

        private const int ExpectedShapesSlideNoTestOneFixed = 5;
        private const int ExpectedShapesSlideNoTestOneDynamic = 8;
        private const int ExpectedShapesSlideNoTestMultipleFixed = 11;
        private const int ExpectedShapesSlideNoTestMultipleDynamic = 14;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDuplicateAndRotate.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_PositionsLabDuplcateAndRotateTest()
        {
            PpOperations.MaximizeWindow();
            var positionsLab = PplFeatures.PositionsLab;
            positionsLab.OpenPane();

            TestOneShapeFixed(positionsLab);
            TestOneShapeDynamic(positionsLab);
            TestMultipleShapesDynamic(positionsLab);
            TestMultipleShapesDynamic(positionsLab);
        }

        private void TestOneShapeFixed(IPositionsLabController positionsLab)
        {
            var actualSlide = PpOperations.SelectSlide(OriginalShapesSlideNoTestOneFixed);

            _shapeNames = new List<string> { Rectangle, Oval };
            Shape shapeStart = PpOperations.SelectShape(Oval)[1];
            Shape shapeEnd = PpOperations.SelectShape(Target)[1];
            PpOperations.SelectShapes(_shapeNames);

            positionsLab.ReorientFixed();
            positionsLab.ToggleDuplicateAndRotateButton();

            RotateShape(shapeStart, shapeEnd);

            var expSlide = PpOperations.SelectSlide(ExpectedShapesSlideNoTestOneFixed);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestOneShapeDynamic(IPositionsLabController positionsLab)
        {
            var actualSlide = PpOperations.SelectSlide(OriginalShapesSlideNoTestOneDynamic);

            _shapeNames = new List<string> { Rectangle, Triangle };
            Shape shapeStart = PpOperations.SelectShape(Triangle)[1];
            Shape shapeEnd = PpOperations.SelectShape(Target)[1];
            PpOperations.SelectShapes(_shapeNames);

            positionsLab.ReorientDynamic();
            positionsLab.ToggleDuplicateAndRotateButton();

            RotateShape(shapeStart, shapeEnd);

            var expSlide = PpOperations.SelectSlide(ExpectedShapesSlideNoTestOneDynamic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestMultipleShapesFixed(IPositionsLabController positionsLab)
        {
            var actualSlide = PpOperations.SelectSlide(OriginalShapesSlideNoTestMultipleFixed);

            _shapeNames = new List<string> { Rectangle, Oval, Triangle };
            Shape shapeStart = PpOperations.SelectShape(Oval)[1];
            Shape shapeEnd = PpOperations.SelectShape(Target)[1];
            PpOperations.SelectShapes(_shapeNames);

            positionsLab.ReorientFixed();
            positionsLab.ToggleDuplicateAndRotateButton();

            RotateShape(shapeStart, shapeEnd);

            var expSlide = PpOperations.SelectSlide(ExpectedShapesSlideNoTestMultipleFixed);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestMultipleShapesDynamic(IPositionsLabController positionsLab)
        {
            var actualSlide = PpOperations.SelectSlide(OriginalShapesSlideNoTestMultipleDynamic);

            _shapeNames = new List<string> { Rectangle, Oval, Triangle };
            Shape shapeStart = PpOperations.SelectShape(Oval)[1];
            Shape shapeEnd = PpOperations.SelectShape(Target)[1];
            PpOperations.SelectShapes(_shapeNames);

            positionsLab.ReorientDynamic();
            positionsLab.ToggleDuplicateAndRotateButton();

            RotateShape(shapeStart, shapeEnd);

            var expSlide = PpOperations.SelectSlide(ExpectedShapesSlideNoTestMultipleDynamic);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        # region Helper methods
        // mouse drag & drop from Control to Shape to apply color
        private void RotateShape(Shape from, Shape to)
        {
            var startPt = new Point(
                PpOperations.PointsToScreenPixelsX(from.Left + from.Width / 2),
                PpOperations.PointsToScreenPixelsY(from.Top + from.Height / 2));
            var endPt = new Point(
                PpOperations.PointsToScreenPixelsX(to.Left + to.Width/2),
                PpOperations.PointsToScreenPixelsY(to.Top + to.Height/2));
            DragAndDrop(startPt, endPt);

            //Need to click away to end rotate
            MouseUtil.SendMouseLeftClick(0, 0);
        }

        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown(startPt.X, startPt.Y);
            MouseUtil.SendMouseUp(endPt.X, endPt.Y);
        }

        private void Click(Control target)
        {
            var pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
        }
        # endregion
    }
}
