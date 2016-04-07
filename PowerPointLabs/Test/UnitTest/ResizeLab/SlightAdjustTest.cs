using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class SlightAdjustTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string ShapeName = "shape";
        private const string ImageName = "image";

        private const int SlightIncreaseWidthSlideNo = 43;
        private const int SlightIncreaseWidthAspectRatioSlideNo = 44;
        private const int SlightDecreaseWidthSlideNo = 45;
        private const int SlightDecreaseWidthAspectRatioSlideNo = 46;
        private const int SlightIncreaseHeightSlideNo = 47;
        private const int SlightIncreaseHeightAspectRatioSlideNo = 48;
        private const int SlightDecreaseHeightSlideNo = 49;
        private const int SlightDecreaseHeightAspectRatioSlideNo = 50;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {ShapeName, ImageName};
            InitOriginalShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightIncreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightIncreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightDecreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightDecreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightIncreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightIncreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightDecreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightDecreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
