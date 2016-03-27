using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabAdjoinTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 4;
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string RotatedRectangle = "Rectangle 6";

        //Results of Operations
        private const int AdjoinHorizontalWithoutAlignNo = 5;
        private const int AdjoinVerticalWithoutAlignNo = 6;
        private const int AdjoinHorizontalWithoutAlignWithRotatedRef = 7;
        private const int AdjoinVerticalWithoutAlignWithRotatedRef = 8;

        private const int AdjoinHorizontalWithAlignNo = 10;
        private const int AdjoinVerticalWithAlignNo = 11;
        private const int AdjoinHorizontalWithAlignWithRotatedRef = 12;
        private const int AdjoinVerticalWithAlignWithRotatedRef = 13;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabAdjoin.pptx";
        }

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
        public void TestAdjoinHorizontalWithoutAlign()
        {
            PositionsLabMain.AdjoinWithoutAligning();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinHorizontalWithoutAlignNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinVerticalWithoutAlign()
        {
            PositionsLabMain.AdjoinWithoutAligning();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinVerticalWithoutAlignNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinHorizontalWithoutAlignWithRotatedRef()
        {
            PositionsLabMain.AdjoinWithoutAligning();
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinHorizontalWithoutAlignWithRotatedRef);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinVerticalWithoutAlignWithRotatedRef()
        {
            PositionsLabMain.AdjoinWithoutAligning();
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinVerticalWithoutAlignWithRotatedRef);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinHorizontalWithAlign()
        {
            PositionsLabMain.AdjoinWithAligning();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinHorizontalWithAlignNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinVerticalWithAlign()
        {
            PositionsLabMain.AdjoinWithAligning();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinVerticalWithAlignNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinHorizontalWithAlignWithRotatedRef()
        {
            PositionsLabMain.AdjoinWithAligning();
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinHorizontalWithAlignWithRotatedRef);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjoinVerticalWithAlignWithRotatedRef()
        {
            PositionsLabMain.AdjoinWithAligning();
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AdjoinVerticalWithAlignWithRotatedRef);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

    }
}
