using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class EqualizeTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string RefShapeName = "reference";
        private const string OvalShapeName = "oval";
        private const string ArrowShapeName = "arrow";

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {RefShapeName, OvalShapeName, ArrowShapeName};
            InitOriginalShapes(SlideNo.EqualizeOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.EqualizeOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeVisualWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeVisualWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeActualWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeActualWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeVisualHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeVisualHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeActualHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeActualHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualWidthAndHeight()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeVisualBoth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthAndHeight()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeActualBoth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
