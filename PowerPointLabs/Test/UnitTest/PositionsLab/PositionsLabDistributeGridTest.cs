using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDistributeGridTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string Rect1 = "Rectangle 3";
        private const string Rect2 = "Rectangle 4";
        private const string Oval3 = "Oval 5";
        private const string RoundRect4 = "Rounded Rectangle 6";
        private const string Rect5 = "Rectangle 7";
        private const string Rect6 = "Rectangle 8";
        private const string Oval7 = "Oval 9";
        private const string RoundRect8 = "Rounded Rectangle 10";
        private const string Rect9 = "Rectangle 11";
        private const string Rect10 = "Rectangle 12";

        //Results of Operations
        private const int DistributeGridWithFirstWithCenterSlide = 5;
        private const int DistributeGridWithFirstAndSecondWithCenterSlide = 6;

        private const int DistributeGridWithFirstWithEdgesSlide = 8;
        private const int DistributeGridWithFirstAndSecondWithEdgesSlide = 9;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDistributeGrid.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridlWithFirstWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, 3, 4);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 4);

            PpOperations.SelectSlide(DistributeGridWithFirstWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridlWithFirstAndSecondWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { Rect1, Rect10, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, 3, 4);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 4);

            PpOperations.SelectSlide(DistributeGridWithFirstAndSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridlWithFirstWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, 3, 4);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 4);

            PpOperations.SelectSlide(DistributeGridWithFirstWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridlWithFirstAndSecondWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { Rect1, Rect10, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, 3, 4);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 4);

            PpOperations.SelectSlide(DistributeGridWithFirstAndSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
