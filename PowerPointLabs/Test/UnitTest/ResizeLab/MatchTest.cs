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

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { ShapeName, ImageName };
            InitOriginalShapes(SlideNo.MatchOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.MatchOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestMatchWidth()
        {
            var actualShapes = GetShapes(SlideNo.MatchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.MatchVisualWidth, _shapeNames);

            _resizeLab.MatchWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestMatchHeight()
        {
            var actualShapes = GetShapes(SlideNo.MatchOrigin, _shapeNames);
            var expectedShapes = GetShapes(SlideNo.MatchVisualHeight, _shapeNames);

            _resizeLab.MatchHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
