using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class StretchShrinkTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string RefShapeName= "ref";
        private const string LeftShapeName = "leftOfRef";
        private const string RightShapeName = "rightOfRef";
        private const string OverShapeName = "overRef";

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {RefShapeName, LeftShapeName, RightShapeName, OverShapeName};
            InitOriginalShapes(SlideNo.StretchOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.StretchOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchLeft, _shapeNames);

            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchLeftAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftOuterMost()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchLeftOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchRight, _shapeNames);

            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchRightAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightOuterMost()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchRightOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchTop, _shapeNames);

            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchTopAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopOuterMost()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchTopOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithoutAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchBottom, _shapeNames);

            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithAspectRatio()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchBottomAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomOuterMost()
        {
            var actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.StretchBottomOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
