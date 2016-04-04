using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class AdjustProportionallyTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string RefShapeName = "reference";
        private const string OvalShapeName = "oval";
        private const string CornerRectangleName = "cornerRectangle";

        private const int OriginalShapesSlideNo = 36;
        private const int AdjustWidthProportionallySlideNo = 37;
        private const int AdjustWidthProportionallyAspectRatioSlideNo = 38;
        private const int AdjustHeightProportionallySlideNo = 39;
        private const int AdjustHeightProportionallyAspectRatioSlideNo = 40;

        private readonly List<float> _proportionList = new List<float>()
        {
            1,
            2,
            3
        };

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { RefShapeName, OvalShapeName, CornerRectangleName };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustWidthProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(AdjustWidthProportionallySlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustWidthProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(AdjustWidthProportionallyAspectRatioSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustHeightProportionallyWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(AdjustHeightProportionallySlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustHeightProportionallyWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var expectedShapes = GetShapes(AdjustHeightProportionallyAspectRatioSlideNo, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
