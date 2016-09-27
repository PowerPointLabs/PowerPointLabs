using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabFlipTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapes0DegreesSlideNo = 3;
        private const int OriginalShapes30DegreesSlideNo = 4;
        private const int OriginalShapes60DegreesSlideNo = 5;
        private const int OriginalShapes105DegreesSlideNo = 6;
        private const int OriginalShapes120DegreesSlideNo = 7;

        private const string Cloud5 = "Cloud 5";
        private const string Smile9 = "Smile 9";
        private const string Plus8 = "Plus 8";
        private const string Heart3 = "Heart 3";
        private const string Arrow1 = "Arrow 1";
        private const string Arrow2 = "Arrow 2";
        private const string Triangle5 = "Triangle 5";
        private const string Rectangle1 = "Rectangle 1";
        private const string Picture1 = "Picture 1";
        private const string Picture2 = "Picture 2";

    //Results of Operations
        private const int FlipHorizontal0DegreesSlideNo = 9;
        private const int FlipHorizontal30DegreesSlideNo = 10;
        private const int FlipHorizontal60DegreesSlideNo = 11;
        private const int FlipHorizontal105DegreesSlideNo = 12;
        private const int FlipHorizontal120DegreesSlideNo = 13;

        private const int FlipVertical0DegreesSlideNo = 15;
        private const int FlipVertical30DegreesSlideNo = 16;
        private const int FlipVertical60DegreesSlideNo = 17;
        private const int FlipVertical105DegreesSlideNo = 18;
        private const int FlipVertical120DegreesSlideNo = 19;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabFlip.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();
            _shapeNames = new List<string> { Cloud5, Smile9, Plus8, Heart3, Arrow1, Arrow2,
                                             Triangle5, Rectangle1, Picture1, Picture2 };
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal0Degrees()
        {
            InitOriginalShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
            //_shapeNames = new List<string> { Cloud5, Smile9, Plus8, Heart3, Arrow1, Arrow2,
            //                                 Triangle5, Rectangle1, Picture1, Picture2 };

            var actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal0DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
            //RestoreShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal30Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes30DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal30DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal60Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes60DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal60DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal105Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes105DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal105DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal120Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes120DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal120DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical0Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipVertical(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical0DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical30Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes30DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipVertical(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical30DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical60Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes60DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipVertical(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical60DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical105Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes105DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipVertical(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical105DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical120Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes120DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipVertical(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical120DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
