using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDistributeTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string BorderRectangle = "Rectangle 1";
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string RotatedRectangle = "Rectangle 6";

        //Results of Operations
        private const int DistributeHorizontalWithinSlideWithEdgesSlide = 5;
        private const int DistributeVerticalWithinSlideWithEdgesSlide = 6;
        private const int DistributeCenterWithinSlideWithEdgesSlide = 7;

        private const int DistributeHorizontalWithinFirstWithEdgesSlide = 9;
        private const int DistributeVerticalWithinFirstWithEdgesSlide = 10;
        private const int DistributeCenterWithinFirstWithEdgesSlide = 11;

        private const int DistributeHorizontalWithinFirstAndSecondWithEdgesSlide = 13;
        private const int DistributeVerticalWithinFirstAndSecondWithEdgesSlide = 14;
        private const int DistributeCenterWithinFirstAndSecondWithEdgesSlide = 15;

        private const int DistributeHorizontalWithinCornerWithEdgesSlide = 17;
        private const int DistributeVerticalWithinCornerWithEdgesSlide = 18;
        private const int DistributeCenterWithinCornerWithEdgesSlide = 19;

        private const int DistributeHorizontalWithinSlideWithCenterSlide = 21;
        private const int DistributeVerticalWithinSlideWithCenterSlide = 22;
        private const int DistributeCenterWithinSlideWithCenterSlide = 23;

        private const int DistributeHorizontalWithinFirstWithCenterSlide = 25;
        private const int DistributeVerticalWithinFirstWithCenterSlide = 26;
        private const int DistributeCenterWithinFirstWithCenterSlide = 27;

        private const int DistributeHorizontalWithinFirstAndSecondWithCenterSlide = 29;
        private const int DistributeVerticalWithinFirstAndSecondWithCenterSlide = 30;
        private const int DistributeCenterWithinFirstAndSecondWithCenterSlide = 31;

        private const int DistributeHorizontalWithinCornerWithCenterSlide = 33;
        private const int DistributeVerticalWithinCornerWithCenterSlide = 34;
        private const int DistributeCenterWithinCornerWithCenterSlide = 35;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDistribute.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinSlideWithEdges()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinSlideWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinSlideWithEdges()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinSlideWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinSlideWithEdges()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinSlideWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinFirstWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinFirstWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinFirstWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinFirstWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinFirstWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinFirstWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinFirstAndSecondWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinFirstAndSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinFirstAndSecondWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinFirstAndSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinFirstAndSecondWithEdges()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinFirstAndSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinCornerWithEdges()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinCornerWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinCornerWithEdges()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinCornerWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinCornerWithEdges()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinCornerWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinSlideWithCenter()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinSlideWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinSlideWithCenter()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinSlideWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinSlideWithCenter()
        {
            PositionsLabMain.DistributeReferToSlide();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinSlideWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinFirstWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinFirstWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinFirstWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinFirstWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinFirstWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { BorderRectangle, UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinFirstWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinFirstAndSecondWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinFirstAndSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinFirstAndSecondWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinFirstAndSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinFirstAndSecondWithCenter()
        {
            PositionsLabMain.DistributeReferToFirstTwoShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, RotatedArrow, Oval, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinFirstAndSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeHorizontalWithinCornerWithCenter()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;

            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(DistributeHorizontalWithinCornerWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeVerticalWithinCornerWithCenter()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(DistributeVerticalWithinCornerWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeCenterWithinCornerWithCenter()
        {
            PositionsLabMain.DistributeReferToExtremeShapes();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth, slideHeight);

            PpOperations.SelectSlide(DistributeCenterWithinCornerWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
