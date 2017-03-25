using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.SyncLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class SyncLabPositionTest : BaseSyncLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string RotatedRectangle = "Rectangle 6";

        //Results of Operations
        private const int AlignShapesLeftToSlideNo = 5;
        private const int AlignShapesRightToSlideNo = 6;
        private const int AlignShapesTopToSlideNo = 7;
        private const int AlignShapesBottomToSlideNo = 8;
        private const int AlignShapesHorizontalToSlideNo = 9;
        private const int AlignShapesVerticalToSlideNo = 10;
        private const int AlignShapesCenterToSlideNo = 11;

        private const int AlignShapesLeftToRefShapeNo = 13;
        private const int AlignShapesRightToRefShapeNo = 14;
        private const int AlignShapesTopToRefShapeNo = 15;
        private const int AlignShapesBottomToRefShapeNo = 16;
        private const int AlignShapesHorizontalToRefShapeNo = 17;
        private const int AlignShapesVerticalToRefShapeNo = 18;
        private const int AlignShapesCenterToRefShapeNo = 19;

        private const int AlignOneShapeLeftDefaultNo = 21;
        private const int AlignOneShapeRightDefaultNo = 22;
        private const int AlignOneShapeTopDefaultNo = 23;
        private const int AlignOneShapeBottomDefaultNo = 24;
        private const int AlignOneShapeHorizontalDefaultNo = 25;
        private const int AlignOneShapeVerticalDefaultNo = 26;
        private const int AlignOneShapeCenterDefaultNo = 27;

        private const int AlignShapesLeftDefaultNo = 28;
        private const int AlignShapesRightDefaultNo = 29;
        private const int AlignShapesTopDefaultNo = 30;
        private const int AlignShapesBottomDefaultNo = 31;
        private const int AlignShapesHorizontalDefaultNo = 32;
        private const int AlignShapesVerticalDefaultNo = 33;
        private const int AlignShapesCenterDefaultNo = 34;

        private const int AlignShapesRadialNo = 36;

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignLeftToSlide()
        {
            PositionsLabMain.AlignReferToSlide();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AlignShapesLeftToSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
