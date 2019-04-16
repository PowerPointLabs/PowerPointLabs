using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ResizeLab;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class MainSettingsTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string WithAspectRatioShapeNames = "withAspectRatio";
        private const string WithoutAspectRatioShapeNames = "withoutAspectRatio";
        private const string ImageName = "image";

        private const int AnchorTopLeftSlideNo = 53;
        private const int AnchorTopCenterSlideNo = 54;
        private const int AnchorTopRightSlideNo = 55;
        private const int AnchorMiddleLeftSlideNo = 56;
        private const int AnchorCenterSlideNo = 57;
        private const int AnchorMiddleRightSlideNo = 58;
        private const int AnchorBottomLeftSlideNo = 59;
        private const int AnchorBottomCenterSlideNo = 60;
        private const int AnchorBottomRightSlideNo = 61;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {WithoutAspectRatioShapeNames, WithAspectRatioShapeNames, ImageName};
            InitOriginalShapes(SlideNo.AnchorOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopLeft;
            RestoreShapes(SlideNo.AnchorOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestLockAspectRatio()
        {
            PowerPoint.ShapeRange shapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);

            _resizeLab.ChangeShapesAspectRatio(shapes, true);

            foreach (PowerPoint.Shape shape in shapes)
            {
                Assert.AreEqual(shape.LockAspectRatio, MsoTriState.msoTrue);
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestUnlockAspectRatio()
        {
            PowerPoint.ShapeRange shapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);

            _resizeLab.ChangeShapesAspectRatio(shapes, false);

            foreach (PowerPoint.Shape shape in shapes)
            {
                Assert.AreEqual(shape.LockAspectRatio, MsoTriState.msoFalse);
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorTopLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualTopLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorTopLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualTopLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorTopCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualTopCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopCenter;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorTopCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualTopCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopCenter;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorTopRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualTopRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorTopRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualTopRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.TopRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorMiddleLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualMiddleLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.MiddleLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorMiddleLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualMiddleLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.MiddleLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.Center;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.Center;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorMiddleRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualMiddleRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.MiddleRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorMiddleRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualMiddleRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.MiddleRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorBottomLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualBottomLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorBottomLeft()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualBottomLeft, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomLeft;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorBottomCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualBottomCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomCenter;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorBottomCenter()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualBottomCenter, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomCenter;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeVisualAnchorBottomRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorVisualBottomRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Visual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestEqualizeActualAnchorBottomRight()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.AnchorOrigin, _shapeNames);
            PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.AnchorActualBottomRight, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoFalse;
            _resizeLab.AnchorPointType = ResizeLabMain.AnchorPoint.BottomRight;
            _resizeLab.ResizeType = ResizeLabMain.ResizeBy.Actual;
            _resizeLab.ResizeToSameHeightAndWidth(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
