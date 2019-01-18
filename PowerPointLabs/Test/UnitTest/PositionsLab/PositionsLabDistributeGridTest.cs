using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin5Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Left()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0LeftSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Center()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignCenter;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0CenterSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter4x4Margin0Right()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignRight;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstCenter4x4Margin0RightSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Top()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0TopSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Center()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignCenter;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0CenterSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstCenter6x3Margin0Bottom()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignRight;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstCenter6x3Margin0BtmSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge4x4Margin5()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstEdge4x4Margin5Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge4x4Margin0()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstEdge4x4Margin0Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstEdge6x3Margin0()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.GridMarginTop = 0;
            PositionsLabSettings.GridMarginLeft = 0;
            PositionsLabSettings.GridMarginRight = 0;
            PositionsLabSettings.GridMarginBottom = 0;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3, Rect16 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstEdge6x3Margin0Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondCenter4x4()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondCenter4x4Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondCenter6x3()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondCenter6x3Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondEdge4x4()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 4, 4);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondEdge4x4Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeGridFirstAndSecondEdge6x3()
        {
            PositionsLabSettings.DistributeReference = PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.GridMarginTop = 5;
            PositionsLabSettings.GridMarginLeft = 5;
            PositionsLabSettings.GridMarginRight = 5;
            PositionsLabSettings.GridMarginBottom = 5;
            PositionsLabSettings.DistributeGridAlignment = PositionsLabSettings.GridAlignment.AlignLeft;

            _shapeNames = new List<string> { Rect1, Rect16, Rect2, Oval3, RoundRect4, Rect5, Rect6, Oval7, RoundRect8, Rect9, Rect10, Pic11, Pic12, Pic3 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, int, int> positionsAction = (shapes, rowLength, colLength) => PositionsLabMain.DistributeGrid(shapes, rowLength, colLength);
            ExecutePositionsAction(positionsAction, actualShapes, 6, 3);

            PpOperations.SelectSlide(DistributeGridFirstAndSecondEdge6x3Slide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
