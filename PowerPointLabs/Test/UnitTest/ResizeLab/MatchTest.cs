using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class MatchTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string ShapeName = "shape";
        private const string ImageName = "image";

        private const int OriginalShapesSlideNo = 32;
        private const int MatchWidthSlideNo = 33;
        private const int MatchHeightSlideNo = 34;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { ShapeName, ImageName };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestMatchWidth()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(MatchWidthSlideNo, _shapeNames);

            _resizeLab.MatchWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestMatchHeight()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(MatchHeightSlideNo, _shapeNames);

            _resizeLab.MatchHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
