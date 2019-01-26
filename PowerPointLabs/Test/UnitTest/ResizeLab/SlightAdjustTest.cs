using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class SlightAdjustTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string ShapeName = "shape";
        private const string ImageName = "image";

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {ShapeName, ImageName};
            InitOriginalShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualIncreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualIncreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualDecreaseWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualDecreaseWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualIncreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseVisualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualIncreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightIncreaseActualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualIncreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.IncreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualDecreaseHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseVisualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightVisualDecreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSlightDecreaseActualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.SlightAdjustOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.SlightActualDecreaseHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.DecreaseHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
