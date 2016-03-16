using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class SameDimensionTest : ResizeLabBaseTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private readonly Dictionary<string, ShapeProperties> _originalShapesProperties =
            new Dictionary<string, ShapeProperties>();
        private List<string> _shapeNames;

        private const string RefShapeName = "reference";
        private const string OvalShapeName = "oval";
        private const string ArrowShapeName = "arrow";

        private const int OriginalShapesSlideNo = 10;
        private const int SameWidthSlideNo = 11;
        private const int SameWidthAspectRatioSlideNo = 12;
        private const int SameHeightSlideNo = 13;
        private const int SameHeightAspectRatioSlideNo = 14;
        private const int SameWidthAndHeightSlideNo = 15;

        [TestInitialize]
        public void TestInitialize()
        {
            PpOperations.SelectSlide(OriginalShapesSlideNo);
            _originalShapesProperties.Clear();
            _shapeNames = new List<string> {RefShapeName, OvalShapeName, ArrowShapeName};

            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames, _originalShapesProperties);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            var shapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            RestoreShapes(shapes, _originalShapesProperties);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.ResizeToSameWidth(actualShapes);

            PpOperations.SelectSlide(SameWidthSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.ResizeToSameWidth(actualShapes);

            PpOperations.SelectSlide(SameWidthAspectRatioSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.ResizeToSameHeight(actualShapes);

            PpOperations.SelectSlide(SameHeightSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.ResizeToSameHeight(actualShapes);

            PpOperations.SelectSlide(SameHeightAspectRatioSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSameWidthAndHeight()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);

            PpOperations.SelectSlide(SameWidthAndHeightSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
