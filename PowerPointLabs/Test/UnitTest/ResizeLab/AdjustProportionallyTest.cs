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
        private const string BlackCornerRectangleName = "blackCornerRectangle";
        private const string BlueCornerRectangleName = "blueCornerRectangle";

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
            InitOriginalShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualWidthProportionallyWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustVisualWidthProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualWidthProportionallyWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustActualWidthProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualWidthProportionallyWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustVisualWidthProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualWidthProportionallyWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustActualWidthProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustWidthProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualHeightProportionallyWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustVisualHeightProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualHeightProportionallyWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustActualHeightProportionally, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustVisualHeightProportionallyWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustVisualHeightProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustActualHeightProportionallyWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustProportionallyOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustActualHeightProportionallyAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustHeightProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustAreaProportionally()
        {
            _shapeNames = new List<string> { RefShapeName, BlackCornerRectangleName, BlueCornerRectangleName };
            InitOriginalShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);

            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustAreaProportionallyAutoShape, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustAreaProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
            RestoreShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);
            
            InitOriginalShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
            actualShapes = GetShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
            expectedShapes = GetShapes(SlideNo.AdjustAreaProportionallyFreeform, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustAreaProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
            RestoreShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAdjustAreaProportionallyWithAspectRatio()
        {
            _shapeNames = new List<string> { RefShapeName, BlackCornerRectangleName, BlueCornerRectangleName };
            InitOriginalShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);

            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AdjustAreaProportionallyAutoShapeAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustAreaProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
            RestoreShapes(SlideNo.AdjustAreaProportionallyAutoShapeOrigin, _shapeNames);

            InitOriginalShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
            actualShapes = GetShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
            expectedShapes = GetShapes(SlideNo.AdjustAreaProportionallyFreeformAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.AdjustProportionallyProportionList = _proportionList;
            _resizeLab.AdjustAreaProportionally(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
            RestoreShapes(SlideNo.AdjustAreaProportionallyFreeformOrigin, _shapeNames);
        }
    }
}
