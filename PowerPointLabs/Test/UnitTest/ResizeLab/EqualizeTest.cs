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

        private const int OriginalShapesSlideNo = 17;
        private const int SameWidthSlideNo = 18;
        private const int SameWidthAspectRatioSlideNo = 19;
        private const int SameHeightSlideNo = 20;
        private const int SameHeightAspectRatioSlideNo = 21;
        private const int SameWidthAndHeightSlideNo = 22;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {RefShapeName, OvalShapeName, ArrowShapeName};
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(SameWidthSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(SameWidthAspectRatioSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(SameHeightSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(SameHeightAspectRatioSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthAndHeight()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(SameWidthAndHeightSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
