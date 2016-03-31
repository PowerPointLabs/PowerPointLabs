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

        private const int OriginalShapesSlideNo = 3;
        private const int StretchLeftSlideNo = 4;
        private const int StretchLeftAspectRatioSlideNo = 5;
        private const int StretchLeftOuterMostSlideNo = 6;
        private const int StretchRightSlideNo = 7;
        private const int StretchRightAspectRatioSlideNo = 8;
        private const int StretchRightOuterMostSlideNo = 9;
        private const int StretchTopSlideNo = 10;
        private const int StretchTopAspectRatioSlideNo = 11;
        private const int StretchTopOuterMostSlideNo = 12;
        private const int StretchBottomSlideNo = 13;
        private const int StretchBottomAspectRatioSlideNo = 14;
        private const int StretchBottomOuterMostSlideNo = 15;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {RefShapeName, LeftShapeName, RightShapeName, OverShapeName};
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchLeftSlideNo, _shapeNames);

            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchLeftAspectRatioSlideNo, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftOuterMost()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchLeftOuterMostSlideNo, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchRightSlideNo, _shapeNames);

            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchRightAspectRatioSlideNo, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightOuterMost()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchRightOuterMostSlideNo, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchTopSlideNo, _shapeNames);

            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchTopAspectRatioSlideNo, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopOuterMost()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchTopOuterMostSlideNo, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchBottomSlideNo, _shapeNames);

            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchBottomAspectRatioSlideNo, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomOuterMost()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(StretchBottomOuterMostSlideNo, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
