﻿using System;
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
        private const string Pic3 = "Picture 3";
        private const string Pic11 = "Picture 11";
        private const string Pic12 = "Picture 12";
        private const string Rect16 = "Rectangle 16";

        //Results of Operations
        private const int DistributeGridFirstCenter4x4Margin5Slide = 5;
        private const int DistributeGridFirstCenter4x4Margin0LeftSlide = 6;
        private const int DistributeGridFirstCenter4x4Margin0CenterSlide = 7;
        private const int DistributeGridFirstCenter4x4Margin0RightSlide = 8;
        private const int DistributeGridFirstCenter6x3Margin0TopSlide = 9;
        private const int DistributeGridFirstCenter6x3Margin0CenterSlide = 10;
        private const int DistributeGridFirstCenter6x3Margin0BtmSlide = 11;

        private const int DistributeGridFirstEdge4x4Margin5Slide = 13;
        private const int DistributeGridFirstEdge4x4Margin0Slide = 14;
        private const int DistributeGridFirstEdge6x3Margin0Slide = 15;

        private const int DistributeGridFirstAndSecondCenter4x4Slide = 17;
        private const int DistributeGridFirstAndSecondCenter6x3Slide = 18;

        private const int DistributeGridFirstAndSecondEdge4x4Slide = 20;
        private const int DistributeGridFirstAndSecondEdge6x3Slide = 21;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDistributeGrid.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin5()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(5);
            PositionsLabMain.SetDistributeMarginLeft(5);
            PositionsLabMain.SetDistributeMarginRight(5);
            PositionsLabMain.SetDistributeMarginBottom(5);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignLeft);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin5Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Left()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignLeft);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0LeftSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Center()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignCenter);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0CenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Right()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignRight);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0RightSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Top()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignLeft);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0TopSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Center()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignCenter);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0CenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Bottom()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);
            PositionsLabMain.SetDistributeGridAlignment(PositionsLabMain.GridAlignment.AlignRight);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0BtmSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge4x4Margin5()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.SetDistributeMarginTop(5);
            PositionsLabMain.SetDistributeMarginLeft(5);
            PositionsLabMain.SetDistributeMarginRight(5);
            PositionsLabMain.SetDistributeMarginBottom(5);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstEdge4x4Margin5Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge4x4Margin0()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstEdge4x4Margin0Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge6x3Margin0()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.SetDistributeMarginTop(0);
            PositionsLabMain.SetDistributeMarginLeft(0);
            PositionsLabMain.SetDistributeMarginRight(0);
            PositionsLabMain.SetDistributeMarginBottom(0);

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstEdge6x3Margin0Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondCenter4x4()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondCenter4x4Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondCenter6x3()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondCenter6x3Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondEdge4x4()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondEdge4x4Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondEdge6x3()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 3, 6);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondEdge6x3Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
