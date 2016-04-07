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
        public void TestSameWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthAndHeight()
        {
            var actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.EqualizeBoth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
