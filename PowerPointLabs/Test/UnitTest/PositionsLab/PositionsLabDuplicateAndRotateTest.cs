using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;
using System.Diagnostics;
using Test.Util;
using System.Drawing;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDuplicateAndRotateTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string Rectangle1 = "Rectangle 1";
        private const string Oval2 = "Oval 2";
        private const string IsoscelesTriangle3 = "Isosceles Triangle 3";

        //Results of Operations
        private const int DuplicateAndRotateSingle1SlideNo = 5;
        private const int DuplicateAndRotateSingle2SlideNo = 6;
        private const int DuplicateAndRotateMutiple1SlideNo = 8;
        private const int DuplicateAndRotateMutiple2SlideNo = 9;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDuplicateAndRotate.pptx";
        }

        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown(startPt.X, startPt.Y);
            MouseUtil.SendMouseUp(endPt.X, endPt.Y);
        }

        private void Click(Point thisPt)
        {
            MouseUtil.SendMouseLeftClick(thisPt.X, thisPt.Y);
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { Rectangle1, Oval2, IsoscelesTriangle3 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplicateAndRoateSingle1()
        {
            _shapeNames = new List<string> { Rectangle1, Oval2, IsoscelesTriangle3 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);


            PpOperations.SelectSlide(DuplicateAndRotateSingle1SlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);

            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplcateAndRoateSingle2()
        {
            _shapeNames = new List<string> { Rectangle1, Oval2, IsoscelesTriangle3 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            PpOperations.SelectSlide(DuplicateAndRotateSingle1SlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);

            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplcateAndRoateMultiple1()
        {
            _shapeNames = new List<string> { Rectangle1, Oval2, IsoscelesTriangle3 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            PpOperations.SelectSlide(DuplicateAndRotateSingle1SlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);

            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplcateAndRoateMultiple2()
        {
            _shapeNames = new List<string> { Rectangle1, Oval2, IsoscelesTriangle3 };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            PpOperations.SelectSlide(DuplicateAndRotateSingle1SlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);

            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }
    }
}
