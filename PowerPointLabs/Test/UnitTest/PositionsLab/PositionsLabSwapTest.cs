using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;
using System.Diagnostics;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabSwapTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string RotatedRectangle = "Rectangle 6";

        //Results of Operations
        private const int SwapLeftToRight1Slide = 5;
        private const int SwapLeftToRight2Slide = 6;
        private const int SwapLeftToRight3Slide = 7;
        private const int SwapLeftToRight4Slide = 8;

        private const int SwapClick1Slide = 10;
        private const int SwapClick2Slide = 11;
        private const int SwapClick3Slide = 12;
        private const int SwapClick4Slide = 13;

        private const int SwapTopLeftSlide = 15;
        private const int SwapTopCenterSlide = 16;
        private const int SwapTopRightSlide = 17;
        private const int SwapMiddleLeftSlide = 18;
        private const int SwapMiddleRightSlide = 19;
        private const int SwapBottomLeftSlide = 20;
        private const int SwapBottomCenterSlide = 21;
        private const int SwapBottomRightSlide = 22;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabSwap.pptx";
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
        public void TestSwapLeftToRight1()
        {
            PositionsLabMain.IsSwapByClickOrder = false;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapLeftToRight1Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapLeftToRight2()
        {
            PositionsLabMain.IsSwapByClickOrder = false;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapLeftToRight2Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapLeftToRight3()
        {
            PositionsLabMain.IsSwapByClickOrder = false;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapLeftToRight3Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapLeftToRight4()
        {
            PositionsLabMain.IsSwapByClickOrder = false;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapLeftToRight4Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapClickOrder1()
        {
            PositionsLabMain.IsSwapByClickOrder = true;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, RotatedArrow, Oval };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapClick1Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapClickOrder2()
        {
            PositionsLabMain.IsSwapByClickOrder = true;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, RotatedArrow, Oval };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapClick2Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]    
        public void TestSwapClickOrder3()
        {
            PositionsLabMain.IsSwapByClickOrder = true;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, RotatedArrow, Oval };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapClick3Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapClickOrder4()
        {
            PositionsLabMain.IsSwapByClickOrder = true;
            PositionsLabMain.SwapReferencePoint = PositionsLabMain.SwapReference.MiddleCenter;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, RotatedArrow, Oval };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);
            ExecutePositionsAction(positionsAction, actualShapes, false);

            PpOperations.SelectSlide(SwapClick4Slide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
