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
        public void TestSlightIncreaseVisualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualIncreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualIncreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualDecreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualDecreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualIncreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualIncreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualDecreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.SlightActualDecreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
