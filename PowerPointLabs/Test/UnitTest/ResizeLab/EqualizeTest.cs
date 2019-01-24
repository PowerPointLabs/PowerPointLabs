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
        public void TestEqualizeVisualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeVisualWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeVisualWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeActualWidth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeActualWidthAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeVisualHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeVisualHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeActualHeight, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeActualHeightAspectRatio, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoTrue;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualWidthAndHeight()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeVisualBoth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualWidthAndHeight()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.EqualizeOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.EqualizeActualBoth, _shapeNames);
            actualShapes.LockAspectRatio = MsoTriState.msoFalse;

            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
