﻿using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class AdjustProportionallyTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string RefShapeName = "reference";
        private const string OvalShapeName = "oval";
        private const string CornerRectangleName = "cornerRectangle";

        private readonly List<float> _proportionList = new List<float>()
        {
            1,
            2,
            3
        };

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { RefShapeName, OvalShapeName, CornerRectangleName };
            InitOriginalShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualWidthProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustVisualWidthProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualWidthProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustActualWidthProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualWidthProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustVisualWidthProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualWidthProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustActualWidthProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualHeightProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustVisualHeightProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualHeightProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustActualHeightProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualHeightProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustVisualHeightProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualHeightProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.AdjustActualHeightProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
