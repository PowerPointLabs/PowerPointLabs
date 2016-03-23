using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
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
        private const int TestStretchLeftSlideNo = 4;
        private const int TestStretchRightSlideNo = 5;
        private const int TestStretchTopSlideNo = 6;
        private const int TestStretchBottomSlideNo = 7;

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
        public void TestStretchLeft()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(TestStretchLeftSlideNo, _shapeNames);

            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRight()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(TestStretchRightSlideNo, _shapeNames);

            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTop()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(TestStretchTopSlideNo, _shapeNames);

            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottom()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(TestStretchBottomSlideNo, _shapeNames);

            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
